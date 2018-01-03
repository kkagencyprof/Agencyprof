VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form shwAdrDetail 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Adressen - AgencyProf"
   ClientHeight    =   6525
   ClientLeft      =   3345
   ClientTop       =   3960
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "shwAdrDetail.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12345
   Begin VB.CheckBox stcky 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   206
      ToolTipText     =   "Keep address in the searchresults of the main form"
      Top             =   6180
      Width           =   255
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   3735
      Left            =   6240
      TabIndex        =   133
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
   Begin MSComctlLib.ListView gd3 
      Height          =   1335
      Left            =   6120
      TabIndex        =   189
      Top             =   3480
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
   Begin VB.CommandButton Command51 
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
      Height          =   200
      Left            =   3720
      TabIndex        =   193
      Top             =   6120
      Width           =   200
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   7
      Left            =   840
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   13
      Text            =   "shwAdrDetail.frx":0BC2
      ToolTipText     =   "Bemerkungen - Doppelklick zur vergrößerten Ansicht"
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox opttel 
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
      Left            =   840
      TabIndex        =   194
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton repert 
      Caption         =   "Repertoire"
      Height          =   195
      Left            =   10080
      TabIndex        =   192
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9240
      Picture         =   "shwAdrDetail.frx":0BC8
      Style           =   1  'Grafisch
      TabIndex        =   191
      ToolTipText     =   "Kalender öffnen"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox ckadrpbez 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8280
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   190
      ToolTipText     =   "Postfach"
      Top             =   3120
      Width           =   855
   End
   Begin MSComctlLib.ListView gd2 
      Height          =   1575
      Left            =   840
      TabIndex        =   187
      ToolTipText     =   "<ctrl>douleclick to create a contact"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
   Begin VB.TextBox cadrpbez 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   188
      ToolTipText     =   "aktuelle Adresse"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "shwAdrDetail.frx":0CC8
      Style           =   1  'Grafisch
      TabIndex        =   185
      ToolTipText     =   "andere Adresse"
      Top             =   1960
      Width           =   255
   End
   Begin VB.CommandButton kh2ja 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Picture         =   "shwAdrDetail.frx":0F3A
      Style           =   1  'Grafisch
      TabIndex        =   184
      ToolTipText     =   "Öffnet die Adressnotizdatei"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton kh2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Picture         =   "shwAdrDetail.frx":115C
      Style           =   1  'Grafisch
      TabIndex        =   183
      ToolTipText     =   "Öffnet die Adressnotizdatei"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Alle zeigen"
      Height          =   255
      Left            =   9960
      TabIndex        =   182
      ToolTipText     =   "Alle Filter löschen, alles anzeigen"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command46 
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
      Left            =   3240
      Picture         =   "shwAdrDetail.frx":137E
      Style           =   1  'Grafisch
      TabIndex        =   181
      ToolTipText     =   "Neue Adresse aus einem Kontakt erstellen"
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox intuse 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   179
      ToolTipText     =   "Postf. geht vor"
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command45 
      Caption         =   "+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9120
      TabIndex        =   178
      ToolTipText     =   "Markierten Kontakt nach oben verschieben"
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command44 
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9360
      TabIndex        =   177
      ToolTipText     =   "Markierten Kontakt nach unten verschieben"
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      MaskColor       =   &H00000000&
      Picture         =   "shwAdrDetail.frx":1F60
      Style           =   1  'Grafisch
      TabIndex        =   176
      ToolTipText     =   "Speichern"
      Top             =   720
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
      Left            =   5520
      TabIndex        =   174
      ToolTipText     =   "Priorität, Änderungen werden sofort gespeichert."
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command43 
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
      Left            =   6600
      Picture         =   "shwAdrDetail.frx":25D2
      Style           =   1  'Grafisch
      TabIndex        =   173
      ToolTipText     =   "Neuen Kontakt aus einer Adresse erstellen"
      Top             =   1080
      Width           =   375
   End
   Begin VB.ListBox List1b 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   4680
      TabIndex        =   172
      ToolTipText     =   "Beziehungen zu anderen Adressen"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command42 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      Picture         =   "shwAdrDetail.frx":31B4
      Style           =   1  'Grafisch
      TabIndex        =   171
      ToolTipText     =   "Beziehungen / Gruppen anzeigen"
      Top             =   840
      Width           =   255
   End
   Begin VB.ListBox impdaten 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   9960
      TabIndex        =   169
      Top             =   10080
      Width           =   3495
   End
   Begin VB.ListBox impfelder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   8400
      TabIndex        =   168
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "shwAdrDetail.frx":3526
      Style           =   1  'Grafisch
      TabIndex        =   167
      ToolTipText     =   "Land aus einer Liste wählen"
      Top             =   3840
      Width           =   135
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "shwAdrDetail.frx":3798
      Style           =   1  'Grafisch
      TabIndex        =   166
      ToolTipText     =   "Land aus einer Liste wählen"
      Top             =   2280
      Width           =   135
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   164
      ToolTipText     =   "Postf. geht vor"
      Top             =   3000
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CheckBox gd1bez 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   162
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Picture         =   "shwAdrDetail.frx":3A0A
      Style           =   1  'Grafisch
      TabIndex        =   161
      ToolTipText     =   "Adresse kopieren"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Picture         =   "shwAdrDetail.frx":3F3C
      Style           =   1  'Grafisch
      TabIndex        =   160
      ToolTipText     =   "Adressdaten den leeren Kontaktfeldern zuweisen"
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00000000&
      Picture         =   "shwAdrDetail.frx":463E
      Style           =   1  'Grafisch
      TabIndex        =   159
      ToolTipText     =   "Alle wählbaren Nummern"
      Top             =   3240
      Width           =   375
   End
   Begin VB.ComboBox season 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9960
      TabIndex        =   158
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox postanredek 
      Height          =   285
      Left            =   6240
      TabIndex        =   156
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox postanredea 
      Height          =   285
      ItemData        =   "shwAdrDetail.frx":47C8
      Left            =   840
      List            =   "shwAdrDetail.frx":47CA
      TabIndex        =   155
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   9000
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   23
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9000
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   21
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6960
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   22
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7800
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   20
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7080
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   19
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox kadat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6960
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   18
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox postf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   7
      ToolTipText     =   "Postfach"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox plzp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   5
      ToolTipText     =   "Postleitzahl des Postfachs"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton icalconf 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Konfig"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Picture         =   "shwAdrDetail.frx":47CC
      Style           =   1  'Grafisch
      TabIndex        =   146
      ToolTipText     =   "Ausgewählte Termine als iCalendar exportieren"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command34 
      Caption         =   "vCard"
      Height          =   375
      Left            =   8520
      TabIndex        =   145
      ToolTipText     =   "Adresse (oder ausgewählten Kontakt) als vCard versenden, Weitergabe der Handynummer einstellbar"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton icalex 
      BackColor       =   &H00C0FFC0&
      Caption         =   "iCAL"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Picture         =   "shwAdrDetail.frx":48CC
      Style           =   1  'Grafisch
      TabIndex        =   144
      ToolTipText     =   "Ausgewählte Termine als iCalendar exportieren"
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Picture         =   "shwAdrDetail.frx":49CC
      Style           =   1  'Grafisch
      TabIndex        =   143
      ToolTipText     =   "Adresse in die Zwischenablage kopieren"
      Top             =   1440
      Width           =   375
   End
   Begin VB.ListBox List10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   141
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "shwAdrDetail.frx":4EFE
      Style           =   1  'Grafisch
      TabIndex        =   140
      ToolTipText     =   "Adresse als Dokument"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "shwAdrDetail.frx":5720
      Style           =   1  'Grafisch
      TabIndex        =   139
      ToolTipText     =   "per Email an Agencyprof"
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Saalplan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   138
      ToolTipText     =   "Neue Biographie anlegen"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Picture         =   "shwAdrDetail.frx":57D2
      Style           =   1  'Grafisch
      TabIndex        =   137
      ToolTipText     =   "... im Explorer öffnen"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Command28 
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
      Left            =   600
      TabIndex        =   136
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox gd1show 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   134
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Bühnenplan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   132
      ToolTipText     =   "Neue Biographie anlegen"
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox datf 
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
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   131
      Text            =   "Text2"
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "shwAdrDetail.frx":5DFC
      Left            =   2640
      List            =   "shwAdrDetail.frx":5DFE
      TabIndex        =   130
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11400
      Sorted          =   -1  'True
      TabIndex        =   129
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   8520
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   16
      ToolTipText     =   "Position"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox Abrede 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "shwAdrDetail.frx":5E00
      Left            =   6960
      List            =   "shwAdrDetail.frx":5E16
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox Anrede 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "shwAdrDetail.frx":5E8D
      Left            =   6960
      List            =   "shwAdrDetail.frx":5E8F
      TabIndex        =   14
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CheckBox usempth 
      Caption         =   "Medienpfad"
      Height          =   255
      Left            =   4680
      TabIndex        =   125
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   285
      ItemData        =   "shwAdrDetail.frx":5E91
      Left            =   11400
      List            =   "shwAdrDetail.frx":5E93
      TabIndex        =   124
      Text            =   "Combo1"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   8520
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      Picture         =   "shwAdrDetail.frx":5E95
      Style           =   1  'Grafisch
      TabIndex        =   78
      ToolTipText     =   "Neuen Termin anlegen"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Caption         =   "neue Bio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   81
      ToolTipText     =   "Neue Biographie anlegen"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      Picture         =   "shwAdrDetail.frx":6227
      Style           =   1  'Grafisch
      TabIndex        =   82
      ToolTipText     =   "Formular schiessen"
      Top             =   5880
      Width           =   255
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Verzeichnis öffnen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   85
      ToolTipText     =   "... im Explorer öffnen"
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ListBox List9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   9960
      TabIndex        =   117
      ToolTipText     =   "weitere Druckvorlagen"
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Picture         =   "shwAdrDetail.frx":6477
      Style           =   1  'Grafisch
      TabIndex        =   99
      ToolTipText     =   "Wiedervorlage"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   1080
      TabIndex        =   97
      Text            =   "Text2"
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox datf 
      Alignment       =   2  'Zentriert
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
      Index           =   14
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox datf 
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
      Index           =   13
      Left            =   1800
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   8160
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   9360
      TabIndex        =   94
      Text            =   "Text12"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   93
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   91
      ToolTipText     =   "Diese Mediadateien sind hinterlegt"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command22 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   90
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   89
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox pnull 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   88
      Top             =   9360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   87
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   86
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton p1cmd 
      Caption         =   "Command19"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   84
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton p1cmd 
      Caption         =   "Command19"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   83
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton p1cmd 
      Caption         =   "Command19"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   80
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton p1cmd 
      Caption         =   "Command19"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   79
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Media Dateien"
      Height          =   375
      Left            =   4680
      TabIndex        =   77
      ToolTipText     =   "Medien anzeigen/ verbergen"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   76
      ToolTipText     =   "Email senden an Kontaktperson"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   74
      ToolTipText     =   "Email senden an Adresse"
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Picture         =   "shwAdrDetail.frx":6AE3
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Memo schreiben"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   70
      Text            =   "730"
      ToolTipText     =   "Wieviel Tage zurück"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11520
      TabIndex        =   69
      Text            =   "365"
      ToolTipText     =   "Wieviel Tage in die Zukunft"
      Top             =   840
      Width           =   375
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   68
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   67
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   9960
      TabIndex        =   66
      ToolTipText     =   "Zusatzfelder je nach gewählten Kategorien"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Go !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11400
      TabIndex        =   65
      ToolTipText     =   "Zeige Verknüpfte Termine"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ListBox List3 
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   9960
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   64
      ToolTipText     =   "Kategorien der Termine"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   9960
      Sorted          =   -1  'True
      TabIndex        =   63
      ToolTipText     =   "Liste verknüpfter Termine"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Zusatz - Infos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   62
      ToolTipText     =   "Zusatz-Informationen anzeigen/ verbergen"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   61
      ToolTipText     =   "Zum Löschen deaktivieren"
      Top             =   240
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Picture         =   "shwAdrDetail.frx":714F
      Style           =   1  'Grafisch
      TabIndex        =   60
      ToolTipText     =   "Löschen gesamten Datensatz"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Abwahl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   59
      ToolTipText     =   "Keine Kontaktperson auswählen"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   6960
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
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
      Left            =   2040
      Picture         =   "shwAdrDetail.frx":763F
      Style           =   1  'Grafisch
      TabIndex        =   57
      ToolTipText     =   "Neue Adresse anlegen"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Kontakt - Historie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   56
      ToolTipText     =   "Bisherige Kontakte zeigen"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   55
      ToolTipText     =   "Kategorie hier entfernen"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5280
      TabIndex        =   54
      ToolTipText     =   "Kategorie hinzufügen"
      Top             =   840
      Width           =   255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   4610
      TabIndex        =   53
      ToolTipText     =   "Hier angewandte Kategorien"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9480
      TabIndex        =   52
      ToolTipText     =   "Zum Löschen deaktivieren"
      Top             =   1680
      Value           =   1  'Aktiviert
      Width           =   135
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "shwAdrDetail.frx":79D1
      Style           =   1  'Grafisch
      TabIndex        =   51
      ToolTipText     =   "Neuen Kontakt anlegen"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Picture         =   "shwAdrDetail.frx":7D63
      Style           =   1  'Grafisch
      TabIndex        =   50
      ToolTipText     =   "Löschen Kontakt"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   6960
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   4560
      TabIndex        =   47
      Text            =   "Text2"
      ToolTipText     =   "Kunden-Nummer anlegen"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaskColor       =   &H00000000&
      Picture         =   "shwAdrDetail.frx":8253
      Style           =   1  'Grafisch
      TabIndex        =   46
      ToolTipText     =   "Speichern"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   6960
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   6960
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   17
      Text            =   "shwAdrDetail.frx":88C5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   8760
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Picture         =   "shwAdrDetail.frx":88CB
      Style           =   1  'Grafisch
      TabIndex        =   39
      ToolTipText     =   "Brief schreiben"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fax"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Picture         =   "shwAdrDetail.frx":8F37
      Style           =   1  'Grafisch
      TabIndex        =   38
      ToolTipText     =   "Fax schreiben"
      Top             =   3240
      Width           =   495
   End
   Begin VB.ListBox idxlist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   10
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   12
      Text            =   "Text2"
      ToolTipText     =   "Internet-Adresse"
      Top             =   5040
      Width           =   3495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   9
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   11
      Text            =   "Text2"
      ToolTipText     =   "Handy-Nummer"
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   6120
      TabIndex        =   33
      Text            =   "Text2"
      ToolTipText     =   "Datum der letzte Änderung"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   5
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   9
      Text            =   "Text2"
      ToolTipText     =   "Fax-Nummer"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   4
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   8
      Text            =   "Text2"
      ToolTipText     =   "Telefon"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   3
      Left            =   840
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      Text            =   "shwAdrDetail.frx":95F3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   2
      Left            =   840
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      Text            =   "shwAdrDetail.frx":95F9
      ToolTipText     =   "Straße/ Postfach"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   1
      Left            =   840
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Text            =   "shwAdrDetail.frx":95FF
      ToolTipText     =   "Namen der Firma"
      Top             =   840
      Width           =   3495
   End
   Begin VB.ComboBox altbvorl 
      Height          =   285
      Left            =   5280
      TabIndex        =   142
      ToolTipText     =   "Alternative Briefvorlage benutzen"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Picture         =   "shwAdrDetail.frx":9605
      Style           =   1  'Grafisch
      TabIndex        =   170
      ToolTipText     =   "Kontaktdaten den leeren Adressfeldern zuweisen"
      Top             =   4560
      Width           =   255
   End
   Begin VB.ListBox klist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   6960
      TabIndex        =   36
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "shwAdrDetail.frx":9D07
      Style           =   1  'Grafisch
      TabIndex        =   186
      ToolTipText     =   "andere Adresse"
      Top             =   3480
      Width           =   195
   End
   Begin VB.CheckBox inclcont 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   202
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11160
      TabIndex        =   204
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox datf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   6
      Left            =   840
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   10
      Text            =   "Text2"
      ToolTipText     =   "Email-Adresse"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ComboBox anumsel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   198
      Text            =   "Combo4"
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox optktel 
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
      Left            =   6960
      TabIndex        =   196
      Top             =   4860
      Width           =   2655
   End
   Begin VB.TextBox kdat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   6960
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox knumsel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   199
      Text            =   "Combo4"
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      Picture         =   "shwAdrDetail.frx":9F79
      Style           =   1  'Grafisch
      TabIndex        =   200
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      Picture         =   "shwAdrDetail.frx":A469
      Style           =   1  'Grafisch
      TabIndex        =   201
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "sticky "
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
      Left            =   2760
      TabIndex        =   207
      ToolTipText     =   "Keep address in the searchresults of the main form"
      Top             =   6180
      Width           =   735
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "extra form"
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
      Left            =   11400
      TabIndex        =   205
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "inkl. Konakte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   203
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Image higrusuch 
      Height          =   240
      Left            =   5000
      Picture         =   "shwAdrDetail.frx":A959
      Stretch         =   -1  'True
      ToolTipText     =   "query"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   197
      Top             =   4860
      Width           =   495
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   195
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   9120
      Picture         =   "shwAdrDetail.frx":AD9B
      ToolTipText     =   "Diese Adresse merken"
      Top             =   3480
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   19
      Left            =   240
      Picture         =   "shwAdrDetail.frx":AF25
      ToolTipText     =   "Diese Adresse merken"
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "vertraulich"
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
      Left            =   1080
      TabIndex        =   180
      Top             =   6180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prio.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   175
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "letzte Änderung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image suchen 
      Height          =   480
      Left            =   240
      Picture         =   "shwAdrDetail.frx":B0AF
      ToolTipText     =   "Zeige in Karte"
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Postf. geht vor"
      Height          =   255
      Left            =   4920
      TabIndex        =   165
      ToolTipText     =   "Haken entfernen, um nur Strasse/Ort zu benutzen"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "immer Bezeichnung zeigen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   163
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "www"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8040
      TabIndex        =   122
      ToolTipText     =   "Internet-Adresse besuchen"
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label idshow 
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
      Height          =   495
      Left            =   120
      TabIndex        =   157
      ToolTipText     =   "Sortiername und eindeutiger Bezeichner der Adresse, Doppelklick zum umbenennen"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Strasse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   154
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "PF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   153
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "PLZP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   152
      ToolTipText     =   "Postleitzahl des Postfachs"
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   151
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "PLZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   150
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Land"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   149
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "PLZPostf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   148
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "(Postf.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   147
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Detailliste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   135
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "PLZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   128
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label kfname 
      Caption         =   "kfname(9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6720
      TabIndex        =   127
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   126
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label kfname 
      Caption         =   "kfname(8)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   123
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   3840
      Picture         =   "shwAdrDetail.frx":B4F1
      ToolTipText     =   "Kontakt löschen verboten"
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Handy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6360
      TabIndex        =   121
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Handy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   120
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   119
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   118
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   1215
      Index           =   3
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image P1 
      Height          =   1215
      Index           =   2
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image P1 
      Height          =   1215
      Index           =   1
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image P1 
      Height          =   1215
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   6360
      Picture         =   "shwAdrDetail.frx":BA15
      Top             =   720
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   0
      Left            =   9120
      Picture         =   "shwAdrDetail.frx":BB00
      ToolTipText     =   "Kontakt löschen verboten"
      Top             =   1560
      Width           =   315
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Weitere Dokumente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9960
      TabIndex        =   116
      ToolTipText     =   "Zusatzfelder je nach gewählten Kategorien"
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Saison/Zeitraum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10080
      TabIndex        =   115
      ToolTipText     =   "Dieser Zeitraum interessiert mich"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "In Kategorien:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9840
      TabIndex        =   114
      ToolTipText     =   "Diese Personengruppen interessieren mich"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "zur."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   112
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Zusatzfelder"
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
      Left            =   9960
      TabIndex        =   111
      ToolTipText     =   "Zusatzfelder je nach gewählten Kategorien"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Kontakte"
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
      Left            =   6960
      TabIndex        =   110
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   109
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "formel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   108
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Anzeigen:"
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
      Left            =   4920
      TabIndex        =   107
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Kategorie 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategorien:"
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
      Left            =   4800
      TabIndex        =   106
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Land"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   105
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Bearbeiten:"
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
      Left            =   4680
      TabIndex        =   104
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   103
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Schluss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   102
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Anrede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   101
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   100
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Bundesland"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   98
      Top             =   9720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label kfname 
      Caption         =   "tfh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   96
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pfad zur Mediendatei"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   92
      ToolTipText     =   "Hier sind die Mediadateien gespeichert"
      Top             =   8040
      Width           =   4815
   End
   Begin VB.Label kfname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   75
      Top             =   9840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "vor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   72
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Max.Tg."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   71
      Top             =   840
      Width           =   615
   End
   Begin VB.Label kfname 
      Caption         =   "kfname(6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   58
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label kfname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   49
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Kürzel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   48
      Top             =   0
      Width           =   615
   End
   Begin VB.Label kfname 
      Caption         =   "kfname(3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   45
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label kfname 
      Caption         =   "kfname(2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   44
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label kfname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   41
      Top             =   9360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label kfname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   40
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "www"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      ToolTipText     =   "Internet-Adresse besuchen"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Notizen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Strasse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   4560
      Shape           =   4  'Gerundetes Rechteck
      Top             =   600
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6015
      Left            =   6120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5775
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   600
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   4560
      Shape           =   4  'Gerundetes Rechteck
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   4560
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   8040
      Shape           =   4  'Gerundetes Rechteck
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Termine:"
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
      Left            =   9840
      TabIndex        =   113
      ToolTipText     =   "Beteiligt an welchen Auftrittsterminen?"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   6375
      Left            =   9840
      Shape           =   4  'Gerundetes Rechteck
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu adr 
      Caption         =   "&Adresse"
      Begin VB.Menu adr_neu 
         Caption         =   "&neu"
         Shortcut        =   ^N
      End
      Begin VB.Menu adr_opn 
         Caption         =   "&Verzeichnis öffnen"
         Shortcut        =   ^O
      End
      Begin VB.Menu adr_sav 
         Caption         =   "&speichern"
         Shortcut        =   ^S
      End
      Begin VB.Menu adr_cop 
         Caption         =   "&kopieren"
         Shortcut        =   ^K
      End
      Begin VB.Menu adr_del 
         Caption         =   "&löschen"
         Enabled         =   0   'False
      End
      Begin VB.Menu adr_ren 
         Caption         =   "&umbenennen"
         Shortcut        =   ^U
      End
      Begin VB.Menu adr_delallow 
         Caption         =   "löschen &erlauben"
         Shortcut        =   ^E
      End
      Begin VB.Menu ruler 
         Caption         =   "----------"
      End
      Begin VB.Menu adr_wvk 
         Caption         =   "&Wiedervorlage"
         Shortcut        =   ^W
      End
      Begin VB.Menu adr_zus 
         Caption         =   "&Zusatz-Infos"
         Shortcut        =   ^Z
      End
      Begin VB.Menu adr_med 
         Caption         =   "&Media-Dateien"
         Shortcut        =   ^D
      End
      Begin VB.Menu ruler2 
         Caption         =   "----------"
      End
      Begin VB.Menu adr_x 
         Caption         =   "&schließen"
      End
   End
   Begin VB.Menu knt 
      Caption         =   "&Kontakte"
      Begin VB.Menu knt_neu 
         Caption         =   "&neu"
      End
      Begin VB.Menu knt_sav 
         Caption         =   "&speichern"
      End
      Begin VB.Menu knt_del 
         Caption         =   "&löschen"
         Enabled         =   0   'False
      End
      Begin VB.Menu knt_delallow 
         Caption         =   "löschen &erlauben"
      End
      Begin VB.Menu knt_dsl 
         Caption         =   "&Auswahl aufheben"
      End
      Begin VB.Menu knt_up 
         Caption         =   "Markierten Kontakt nach &oben"
         Enabled         =   0   'False
         Shortcut        =   +{F12}
      End
      Begin VB.Menu knt_dwn 
         Caption         =   "Markierten Kontakt nach &unten"
         Enabled         =   0   'False
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu dok 
      Caption         =   "&Dokumente u. Aktionen"
      Begin VB.Menu dok_fax 
         Caption         =   "&Fax"
         Shortcut        =   ^F
      End
      Begin VB.Menu dok_brief 
         Caption         =   "&Brief"
         Shortcut        =   ^B
      End
      Begin VB.Menu dok_memo 
         Caption         =   "&Memo"
         Shortcut        =   ^M
      End
      Begin VB.Menu dok_eml 
         Caption         =   "&Email"
         Shortcut        =   ^L
      End
      Begin VB.Menu dok_kont 
         Caption         =   "&Kontakthistorie"
         Shortcut        =   ^T
      End
      Begin VB.Menu ruler3 
         Caption         =   "----------"
      End
      Begin VB.Menu adr_tel 
         Caption         =   "&Telefonanwahl"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu kat 
      Caption         =   "&Kategorien"
      Begin VB.Menu kat_add 
         Caption         =   "&Adresskategorie hinzufügen"
         Shortcut        =   ^H
      End
      Begin VB.Menu kat_del 
         Caption         =   "Adresskategorie en&tfernen"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "&?"
      Begin VB.Menu hlp_hlp 
         Caption         =   "&Hilfe"
      End
   End
End
Attribute VB_Name = "shwAdrDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fl_rl3%, esccnt As Integer, esccnt2 As Integer, esckcnt As Integer
Dim nflds As Integer, nfldsk As Integer, nfldska As Integer, prv$
Dim rcarr$(200), p1offs%, p1cmdmax%, break%, adrnotz$
Dim toffsvon As Long, toffsbis As Long, dirty%
Dim wd0, wd1, hg0, hg1, nl3fl%, gd1upd%
Public srchit%, rlist2icalmode As Boolean, l1bdont As Boolean
Dim honx$(0 To 99), honxptr%, mytopmerk As Integer
Dim usekpos As Boolean, nodbupd As Boolean, autokategorie$
Dim katnames$(99), katid%(99), c4no As Boolean, stcky_igno As Boolean

Private Sub svkadr(bez As String)
Dim c$, cid$

cid$ = datf(0).text
If cid$ = "" Then Exit Sub
kid$ = kdat(0).text
If kid$ = "" Then Exit Sub

c$ = "delete from opt_adresspool where vid='" + cid$ + "' and kid='" + kid$ + "' and Beschreibung='" + trm(bez) + "'"
Call form1.sqlqry(c$)
c$ = "insert into opt_adresspool (id,vid,kid,Beschreibung,Ort,Strasse,PLZ,Postfach,Land,Bundesland,PLZPostfach) values('" + cid$ + kid$ + bez + "','" + cid$ + "','" + kid$ + "',"
c$ = c$ + "'" + trm(bez) + "',"
c$ = c$ + "'" + trm(kadat(3).text) + "',"
c$ = c$ + "'" + trm(kadat(0).text) + "',"
c$ = c$ + "'" + trm(kadat(2).text) + "',"
c$ = c$ + "'" + trm(kadat(5).text) + "',"
c$ = c$ + "'" + trm(kadat(1).text) + "','',"
c$ = c$ + "'" + trm(kadat(4).text) + "')"
Call form1.sqlqry(c$)
End Sub

Private Sub svadr(bez As String)
Dim c$, cid$

cid$ = datf(0).text
If cid$ = "" Then Exit Sub

c$ = "delete from opt_adresspool where vid='" + cid$ + "' and kid='-1' and Beschreibung='" + trm(bez) + "'"
Call form1.sqlqry(c$)
c$ = "insert into opt_adresspool (id,vid,kid,Beschreibung,Ort,Strasse,PLZ,Postfach,Land,Bundesland,PLZPostfach) values('" + cid$ + "-1" + bez + "','" + cid$ + "','-1',"
c$ = c$ + "'" + trm(bez) + "',"
c$ = c$ + "'" + trm(datf(3).text) + "',"
c$ = c$ + "'" + trm(datf(2).text) + "',"
c$ = c$ + "'" + trm(datf(13).text) + "',"
c$ = c$ + "'" + trm(postf.text) + "',"
c$ = c$ + "'" + trm(datf(14).text) + "','',"
c$ = c$ + "'" + trm(plzp.text) + "')"
Call form1.sqlqry(c$)
End Sub

Sub rlist2()
Dim hon$, lvitem, cat$, dtt$, dtg$, knam$, va$
Dim rtmp As ADODB.Recordset, c$, prvrtmpid As String, fkritp%, kid$, nosel, dv$, gnr$, halle$
Dim rs As ADODB.Recordset, wrkl As ADODB.Recordset, tm As Long, prvt$, rtmpid$, db$, cmd$, cmx$, fld, honfn$
Dim stmp As ADODB.Recordset, rrr, slst As ADODB.Recordset, vlst As ADODB.Recordset, honi%, fe$, xld$, wrkn$
Dim seli%, i%, j%, fnam$, ical$, icalo%, hin%, hiw$, auchadr As String, o%

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "rlist2"
xld$ = form1.getusersetting("exceldelimiter", ",")
If Command14.Caption = transe("Zusatz-Infos") Then Exit Sub
fkritp% = 0
List2.Clear
gd1.ListItems.Clear
Call form1.dbg2f("rlist2 start")
kid$ = trm(datf(0).text)
If klist.ListIndex >= 0 Then
  kname$ = kdat(2).text
  kid$ = kdat(2).text + " {" + kid$ + "}"
End If
auchadr = form1.getusersetting("terminlisteauchadresse", "ja")
inclcont.value = 0: If auchadr = "ja" Then inclcont.value = 1
'If kid$ = "" Then Exit Sub
MousePointer = 11: DoEvents
icalo% = -1
nosel = 1
dv$ = datum2sql(Date - toffsvon)
db$ = datum2sql(Date + toffsbis)
For i% = 0 To List3.ListCount - 1
  If List3.Selected(i%) = True Then
    i% = List3.ListCount
    nosel = 0
  End If
Next i%
seli% = -1
If rlist2icalmode Then
  ical$ = form1.s0dir() & "\" & form1.medien() & "\" & kid$ & ".ics"
  icalo% = FreeFile
  Open ical$ For Append As #icalo%
End If

Set slst = New ADODB.Recordset
slst.CursorLocation = adUseServer
'c$ = "SELECT auftritthigru.auftrittstyp FROM auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid " & _
'     "Where (((auftritthigru.felddaten) = '" & kid$ & "') And ((auftritt.Datum) >= '" & dv$ & "') And ((auftritt.Datum) <= '" & db$ & "')) " & _
'     "ORDER BY auftritthigru.auftrittstyp, auftritt.id;"
c$ = "SELECT auftritthigru.auftrittstyp,auftritthigru.auftrittsid FROM auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid "
If kid$ <> "" Then
  If kname$ <> "" Then
    c$ = c$ + "Where (auftritthigru.felddaten='" & kid$ & "' or auftritthigru.felddaten='" & kname$ & "'"
    If auchadr = "ja" Then c$ = c$ + " or auftritthigru.felddaten='" & cut_d1(cut_d2bis(kid$, "{"), "}") & "'"
    c$ = c$ + ") and "
  Else
    c$ = c$ + "Where (auftritthigru.felddaten='" & kid$ & "'"
    If auchadr = "ja" Then c$ = c$ + " or auftritthigru.felddaten like '%{" & kid$ & "}'"
    c$ = c$ + ") and "
  End If
  c$ = c$ + "auftritt.Datum >= '" & dv$ & "' And auftritt.Datum <= '" & db$ & "' "
' zu unscharf: "instr(lcase(auftritthigru.felddaten),'{" & LCase(kid$) & "}')>0 ) And ((auftritt.Datum) >= '" & dv$ & "') And ((auftritt.Datum) <= '" & db$ & "')) "
Else
  c$ = c$ + "Where (((auftritt.Datum) >= '" & dv$ & "') And ((auftritt.Datum) <= '" & db$ & "')) "

End If
c$ = c$ + "ORDER BY auftritthigru.auftrittstyp, auftritt.id;"
Call tm_start(0)
Set slst = New ADODB.Recordset
slst.CursorLocation = adUseServer
Debug.Print c$
rrr = form1.adoopen(slst, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Call form1.dbg2f("Fehler " + trm(rrr) + " " + Error$(rrr))
  MousePointer = 0
  Exit Sub
End If
tm = tm_stop(0)
prvt$ = ""
While Not slst.EOF
'Debug.Print trm(slst!auftrittsid) + " " + slst!auftrittstyp
If prvt$ <> slst!auftrittstyp Then
prvt$ = slst!auftrittstyp
If Combo1.text = "CSV-Export" Or Combo1.text = "GEMA-Export" Then
  o% = FreeFile
  Open form1.mydir() + "\_Termine.csv" For Append As #o%
End If
For i% = 0 To List3.ListCount - 1
  If prvt$ = transo(List3.List(i%)) Then
  If nosel = 1 Or List3.Selected(i%) Then

  cmd$ = "SELECT usr_" & utabn(transo(List3.List(i%))) & ".*, auftritt.id as aid,auftritt.bezeichnung," & _
       "auftritt.ort as ort,auftritt.datum as adatum ,auftritt.astatus as astatus "
  cmd$ = cmd$ + "FROM " & _
              " (auftritt INNER JOIN usr_" & utabn(transo(List3.List(i%))) & " ON auftritt.id = usr_" & utabn(transo(List3.List(i%))) & ".id)" + _
              " INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid"

  cmx$ = ""
  honi% = 0
  hin% = 0
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseServer
  c$ = "select * from usr_" & utabn(transo(List3.List(i%))) & " where id='nw'"
  Call tm_start(0)
rrr = form1.adoopen(rs, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  j% = 0: honxptr% = 0
  For Each fld In rs.Fields
    fnam$ = LCase(fld.name)
    If InStr(LCase(fnam$), "honorar") = 1 Then
      honx$(honxptr%) = LCase(fnam$) & "::" & trm(j%)
      honxptr% = honxptr% + 1
    End If
    If honi% = 0 And fnam$ = "honorar" Then honi% = j%
    If hin% = 0 Then
      If Left(LCase(transo(fnam$)), 7) = "hinweis" Or _
        LCase(transo(fnam$)) = "nachricht" Or _
        LCase(transo(fnam$)) = "tätigkeit" Or _
        Left(LCase(transo(fnam$)), 8) = "anmerkun" Or _
        LCase(transo(fnam$)) = "bemerkung" Then
        hin% = j%
      End If
    End If
    If kid$ <> "" Then
      fe$ = form1.isofadr(LCase(transo(List3.List(i%))), fnam$)
      If fe$ <> "" And kid$ <> "" Then
        If Len(cmx$) > 0 Then
          cmx$ = cmx$ + " or "
        Else
          cmx$ = "( "
        End If
        cmx$ = cmx$ & " (usr_" & utabn(transo(List3.List(i%))) & "." & fnam$ & "='" & kid$ & "')"
      End If
    End If
    j% = j% + 1
  Next fld

'zu unscharf: cmx$ = " WHERE (( (instr(lcase(auftritthigru.FeldDaten),'" + LCase(kid$) + "')>0)"
  If kid$ <> "" Then
    If kname$ = "" Then
      cmx$ = " WHERE (lcase(auftritthigru.FeldDaten)='" + LCase(kid$) + "'"
      If auchadr = "ja" Then
        'cmx$ = cmx$ + " or instr(auftritthigru.FeldDaten,'{" + LCase(kid$) + "}')>0"
        cmx$ = cmx$ + " or lcase(auftritthigru.FeldDaten) like '%{" + LCase(kid$) + "}'"
      End If
    Else
      cmx$ = " WHERE ( lcase(auftritthigru.FeldDaten)='" + LCase(kid$) + "' or lcase(auftritthigru.FeldDaten)='" + LCase(kname$) + "' "
      If auchadr = "ja" Then
        'cmx$ = cmx$ + " or instr(auftritthigru.FeldDaten,'{" + LCase(kid$) + "}')>0"
        cmx$ = cmx$ + " or auftritthigru.FeldDaten='" + LCase(cut_d1(cut_d2bis(kid$, "{"), "}")) + "'"
      End If
    End If
  Else
    cmx$ = " WHERE ( 1 "
  End If
  tm = tm_stop(0)
  Call form1.dbg2f("Timer 0 (fld-loop)=" & trm(tm))

  If Len(cmx$) > 0 Then cmx$ = cmx$ + " ) and "
  cmd$ = cmd$ & cmx$ & "auftritt.datum>='" + dv$ + "' and auftritt.datum<='" + db$ + "'"
  'cmd$ = cmd$ + " ORDER BY auftritt.Datum DESC"
  cmd$ = cmd$ + " ORDER BY auftritt.id DESC"
 Debug.Print cmd$
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  'On Error Resume Next
  Call tm_start(1)
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  tm = tm_stop(1)
  'rrr = Err
  rrr = 0
  'On Error GoTo 0
  If rrr = 0 Then
  rtmpid$ = ""
  While Not rtmp.EOF
    rtmpid$ = rtmpid$ + rtmp!id + "|"
    rtmp.MoveNext
  Wend
  On Error Resume Next
  rtmp.MoveFirst
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
  If Not rtmp.EOF Then
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
    Call tm_start(2)
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & rtmp!id & "' and felddaten='" & kid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    tm = tm_stop(2)
    Call form1.dbg2f("msTimer=" & trm(tm))
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      Set rs = New ADODB.Recordset
      rs.CursorLocation = adUseServer
      Call tm_start(3)
      For j% = 0 To honxptr% - 1
        If InStr(honx$(j%), "honorar" & LCase(honfn$)) = 1 Then
          honi% = InStr(honx$(j%), "::")
          honi% = Val(Mid$(honx$(j%), honi% + 2))
          Exit For
        End If
      Next j%
    End If
  End If
  End If
  prvrtmpid = ""
  While Not rtmp.EOF
    If trm(rtmp!id) <> prvrtmpid Then
    prvrtmpid = trm(rtmp!id)
    hiw$ = ""
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
'    Call tm_start(4)
rrr = form1.adoopen(stmp, "select * from auftritt where id='" & rtmp!id & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
'    tm = tm_stop(4)
'    Call form1.dbg2f("msTimer=" & trm(tm))
    If Not stmp.EOF Then
      DoEvents
      hon$ = ""
      If honi% > 0 Then
        hon$ = trm(rtmp.Fields(honi%).value)
      End If
      hiw$ = ""
      If hin% > 0 Then
        hiw$ = trm(rtmp.Fields(hin%).value)
      End If
      If trm(hiw$) = "" Or gd1bez.value = 1 Then hiw$ = trm(stmp!bezeichnung)
      If Combo1.text = "Cloud-Export" Then
        form1.hordexlock = True
        Call form1.event2cloud(rtmp!id)
        form1.hordexlock = False
      End If
      If Combo1.text = "CSV-Export" Then
        Print #o%, """" + stmp!datum + """" + xld$ + """" + transe(stmp!auftrittstyp) + """" + xld$ + """" + trm(stmp!bezeichnung) + """" + xld$ + """" + trm(stmp!ort) + """";
        For j% = 1 To rtmp.Fields.Count - 1
          If rtmp.Fields(j%).name <> "aid" Then
            Print #o%, xld$ + """" + trm(rtmp.Fields(j%).value) + """";
          End If
        Next j%
        Print #o%,
      End If
      If Combo1.text = "GEMA-Export" Then
        cmd$ = "select FeldDaten as wert from auftritthigru where FeldName='Programm' and auftrittsid='" + trm(rtmp!id) + "'"
        c$ = form1.get1erg(cmd$)
        If c$ <> "" Then
          cmd$ = "select werkid from programmliste where programmid='" + c$ + "'"
          Set wrkl = New ADODB.Recordset
          wrkl.CursorLocation = adUseServer
          rrr = form1.adoopen(wrkl, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If rrr = 0 Then
            wrkn$ = ""
'Verlag und Album raus 171126
            While Not wrkl.EOF
              cmd$ = "select s14 as wert from w_loc where id='" + trm(wrkl!werkid) + "'"
              gnr$ = trm(form1.get1erg(cmd$))
'171126
'              cmd$ = "select s13 as wert from w_loc where id='" + trm(wrkl!werkid) + "'"
'              tntr$ = trm(form1.get1erg(cmd$))
              If wrkn$ <> "" And LCase(wrkn$) <> "pause" Then wrkn$ = wrkn$ + vbCrLf
              kn$ = form1.getkompnamebywerkid(trm(wrkl!werkid))
              If kn$ <> "" And LCase(wrkn$) <> "pause" Then wrkn$ = wrkn$ + kn$ + ": " + form1.getwerknamebyid(trm(wrkl!werkid)) + " GEMA: " + gnr
'171126
'              If tntr$ <> "" Then wrkn$ = wrkn$ + vbCrLf + "Album: " + tntr$
'              If Not form1.isfieldmissing("opt_published", "id") Then
'                cmd$ = "select aid from opt_published where wid='" + trm(wrkl!werkid) + "'"
'                Set vlst = New ADODB.Recordset
'                vlst.CursorLocation = adUseServer
'                rrr = form1.adoopen(vlst, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
'                If rrr = 0 Then
'                  If Not vlst.EOF Then wrkn$ = wrkn$ + vbCrLf + "Verlag:"
'                  While Not vlst.EOF
'                    wrkn$ = wrkn$ + vbCrLf + strrepl(form1.getnamebyid(vlst!aid), vbCrLf, " - ")
'                    vlst.MoveNext
'                  Wend
'                End If
'              End If
              wrkl.MoveNext
            Wend
            If wrkn <> "" Then
              cmd$ = "select FeldDaten as wert from auftritthigru where (FeldName='Veranstalter' or FeldName='Medium') and auftrittsid='" + trm(rtmp!id) + "'"
              va$ = trm(form1.get1erg(cmd$)): va$ = form1.getAdrProperty(va$, "Name")
              cmd$ = "select FeldDaten as wert from auftritthigru where (FeldName='Halle' or FeldName='Saal') and auftrittsid='" + trm(rtmp!id) + "'"
              halle$ = trm(form1.get1erg(cmd$)): halle$ = form1.getAdrProperty(halle$, "Name")
              Print #o%, """" + stmp!datum + """" + xld$ + """" + transe(stmp!auftrittstyp) + """" + xld$ + """" + trm(stmp!ort) + vbCrLf + halle$ + """";
              Print #o%, xld$ + """" + va$ + """";
              Print #o%, xld$ + """" + wrkn$ + """"
            End If
          End If
        End If
      End If
      List2.AddItem stmp!datum & " " & transe(stmp!auftrittstyp) & "(" & stmp!bezeichnung & ")" & transe(" in ") & stmp!ort & Space$(80) + "(AID:" & rtmp!id
      Set lvitem = gd1.ListItems.add(, , stmp!datum & "(" & form1.dayofweek(stmp!datum) + ") ")
      lvitem.SubItems(1) = form1.get_atabkz(trm(stmp!auftrittstyp))
      lvitem.SubItems(2) = trm(stmp!ort)
      lvitem.SubItems(3) = hon$
      lvitem.SubItems(4) = hiw$
      lvitem.SubItems(5) = trm(stmp!TourneeplanID)
      lvitem.SubItems(6) = rtmp!id
      If rlist2icalmode Then
        Print #icalo%, "BEGIN:VEVENT"
        Print #icalo%, "UID:" & rtmp!id
        Print #icalo%, "SUMMARY:" & strrepl(nouml(stmp!bezeichnung), ",", "")
        Print #icalo%, "DESCRIPTION:" & strrepl(strrepl(nouml(hiw$), ",", "\qwertz/"), "\qwertz/", ",")
        cat$ = trm(stmp!ort)
        If cat$ <> "" Then
          Print #icalo%, "LOCATION:" & cat$
        End If
        cat$ = form1.getusersetting("MOZCAT_" & stmp!auftrittstyp)
        If cat$ = "" Then cat$ = form1.getusersetting("MOZCAT")
        If cat$ = "" Then cat$ = "Miscellaneous"
        Print #icalo%, "CATEGORIES:" & cat$
        cat$ = form1.getusersetting("MOZSTAT_" & form1.get_eventstatusname(stmp!astatus))
        If cat$ = "" Then cat$ = form1.getusersetting("MOZSTAT")
        If cat$ = "" Then cat$ = "Tentative"
        Print #icalo%, "STATUS:" & cat$
        Print #icalo%, "CLASS:PUBLIC"
        dtt$ = Trim(strrepl("" & onlynums(trm(stmp!zeit)), ":", ""))
        While Len(dtt$) < 6: dtt$ = dtt$ + "0": Wend
        dtg$ = strrepl("" & stmp!datum, "-", "") & "T" & dtt$
        Print #icalo%, "DTSTART:" & dtg$
        Print #icalo%, "DTEND:" & dtg$
        dtg$ = strrepl("" & datum2sql(Date), "-", "") & "T" & strrepl("" & Time, ":", "")
        Print #icalo%, "DTSTAMP:" & dtg$
        Print #icalo%, "END:VEVENT"
      End If
    End If
    stmp.Close
    End If
    Label36.Caption = form1.inmylanguage("Termine: ") + trm(List2.ListCount)
    DoEvents
    rtmp.MoveNext
  Wend
Call form1.dbg2f("nach wend; i%=" + trm(i%) + " (* from auftritthigru)")
  rtmp.Close
  End If    'rrr<>0
  End If
  End If
Next i%
If Combo1.text = "CSV-Export" Or Combo1.text = "GEMA-Export" Then
  Close #o%
End If

End If
slst.MoveNext
Wend
Call form1.dbg2f("nach wend slst")
If rlist2icalmode Then
  Close #icalo%
End If
If Combo1.text = "Cloud-Export" Then
        form1.hordexlock = True
form1.cldpusher.Interval = 1000
        form1.hordexlock = False
End If


On Error Resume Next
Call gd1.SetFocus

exsub:
On Error GoTo 0
MousePointer = 0

End Sub

Private Sub Abrede_Change()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Abrede_Change"
id$ = kdat(0).text
If id$ = "" Then
  Command4.Enabled = True
  adr_sav.Enabled = True
Else
  Command5.Enabled = True
  knt_sav.Enabled = True
End If
BackColor = form1.dirtycolor()

End Sub

Private Sub Abrede_Click()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Abrede_Click"
id$ = kdat(0).text
If id$ = "" Then
  Command4.Enabled = True
  adr_sav.Enabled = True
Else
  Command5.Enabled = True
  knt_sav.Enabled = True
End If
BackColor = form1.dirtycolor()


End Sub

Private Sub Abrede_DropDown()
Dim adt$, nam$, pa$, nn$, vn$, pax$
Dim rtmp As ADODB.Recordset


'd2infile = "shwAdrDetail": d2insub = "Anrede_DropDown"
adt$ = trm(Anrede.text)
Abrede.Clear
If klist.ListIndex >= 0 Then
  nam$ = kdat(2).text
  pa$ = postanredek.text
Else
  nam$ = datf(1).text
  pa$ = postanredea.text
End If
nn$ = word2bis(nam$)
vn$ = word1(nam$)
pax$ = "x"
If InStr(LCase(pa$), "frau") > 0 Or InStr(LCase(pa$), "mrs.") > 0 Then pax$ = "w"
If InStr(LCase(pa$), "herr") > 0 Or InStr(LCase(pa$), "mr.") > 0 Then pax$ = "m"
Call abadd("Mit freundlichen Grüßen")
Call abadd("Mit besten Grüßen")
Call abadd("Beste Grüße")
Call abadd("Mit herzlichen Grüßen")
Call abadd("Herzliche Grüße")
Call abadd("Herzlich grüßend")
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_weitereanrede')>0 or instr(owner,'sysvar_system_weitereabrede')>0", form1.adoc, adOpenDynamic, adLockReadOnly)
While Not rtmp.EOF
  If InStr(rtmp!Owner, "anredew") > 0 Then
    If pax$ = "x" Or pax$ = "w" Then Call abadd(trm(rtmp!wert))
  Else
    If InStr(rtmp!Owner, "anredem") > 0 Then
      If pax$ = "x" Or pax$ = "m" Then Call abadd(trm(rtmp!wert))
    Else
      Call abadd(trm(rtmp!wert))
    End If
  End If
  rtmp.MoveNext
Wend

End Sub

Private Sub adr_cop_Click()
Call Command37_Click
End Sub

Private Sub adr_delallow_Click()
If Check2.value = 0 Then
  Check2.value = 1
Else
  Check2.value = 0
End If
End Sub

Private Sub adr_med_Click()
Call Command18_Click
End Sub

Private Sub adr_neu_Click()
Call Command11_Click
End Sub

Private Sub adr_opn_Click()
Call Command29_Click
End Sub

Private Sub adr_ren_Click()
Call idshow_DblClick
End Sub

Private Sub adr_sav_Click()
Call Command4_Click
End Sub

Private Sub adr_tel_Click()
Call Command35_Click
End Sub

Private Sub adr_wvk_Click()
Call Command25_Click
End Sub

Private Sub adr_x_Click()
Call Command1_Click
End Sub

Private Sub adr_zus_Click()
Call Command14_Click
End Sub

Private Sub altbvorl_Click()
Dim i%

'd2infile = "shwAdrDetail": d2insub = "altbvorl_Click"
i% = altbvorl.ListIndex
If i% < 0 Then Exit Sub
Command3.ToolTipText = "Brief erstellen, Vorlage: " & altbvorl.List(i%)

End Sub

Private Sub altbvorl_DropDown()
Dim rtmp As ADODB.Recordset
Dim o$, uId$, c$, rrr

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "altbvorl_DropDown"
altbvorl.Clear
altbvorl.AddItem form1.meinebriefvorlage()

uId$ = form1.getuserid()
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_briefvorlage')>0", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  o$ = rtmp!Owner
  o$ = Mid$(o$, InStr(o$, uId$) + Len(uId$) + 1)
  altbvorl.AddItem rtmp!wert
  rtmp.MoveNext
Wend
rtmp.Close
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM sysvars where instr(owner,'sysvar_system_briefvorlage')>0"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  o$ = rtmp!Owner
  o$ = Mid$(o$, InStr(o$, uId$) + 7)
  altbvorl.AddItem rtmp!wert
  rtmp.MoveNext
Wend
rtmp.Close
End Sub

Private Sub Anrede_Change()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Anrede_Change"
id$ = kdat(0).text
If id$ = "" Then
  Command4.Enabled = True
  adr_sav.Enabled = True
Else
  Command5.Enabled = True
  knt_sav.Enabled = True
End If
BackColor = form1.dirtycolor()

End Sub

Private Sub Anrede_Click()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Anrede_Click"
id$ = kdat(0).text
If id$ = "" Then
  Command4.Enabled = True
  adr_sav.Enabled = True
Else
  Command5.Enabled = True
  knt_sav.Enabled = True
End If
BackColor = form1.dirtycolor()

End Sub

Private Sub Anrede_DropDown()
Dim adt$, nam$, pa$, nn$, vn$, pax$
Dim rtmp As ADODB.Recordset


'd2infile = "shwAdrDetail": d2insub = "Anrede_DropDown"
adt$ = trm(Anrede.text)
Anrede.Clear
If klist.ListIndex >= 0 Then
  nam$ = kdat(2).text
  pa$ = postanredek.text
Else
  nam$ = datf(1).text
  pa$ = postanredea.text
End If
nn$ = word2bis(nam$)
vn$ = word1(nam$)
pax$ = "x"
If InStr(LCase(pa$), "frau") > 0 Or InStr(LCase(pa$), "mrs.") > 0 Then pax$ = "w"
If InStr(LCase(pa$), "herr") > 0 Or InStr(LCase(pa$), "mr.") > 0 Then pax$ = "m"
If form1.getusersetting("anreden_de", "ja") = "ja" Then
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Sehr geehrte Frau " & nn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Sehr geehrter Herr " & nn$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Liebe " & vn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Lieber " & vn$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Liebe Frau " & nn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Lieber Herr " & nn$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Sehr geehrte Frau " & nam$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Sehr geehrter Herr " & nam$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Liebe " & nam$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Lieber " & nam$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Sehr geehrte Frau")
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Sehr geehrter Herr")
  Call anadd("Sehr geehrte Damen und Herren")
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Liebe")
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Lieber")
End If
If form1.getusersetting("anreden_en", "nein") = "ja" Then
  Call anadd("Dear")
End If
If form1.getusersetting("anreden_fr", "nein") = "ja" Then
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Cher " + vn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Chère " + vn$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Cher Monsieur")
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Chère Madame")
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Cher")
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Chère")
End If
If form1.getusersetting("anreden_fl", "nein") = "ja" Then
  Call anadd("Beste " + vn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Geachte Heer " + nn$)
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Geachte Mevrouw " + nn$)
  If pax$ = "x" Or pax$ = "m" Then Call anadd("Geachte Heer")
  If pax$ = "x" Or pax$ = "w" Then Call anadd("Geachte Mevrouw")
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_weitereanrede')>0 or instr(owner,'sysvar_system_weitereanrede')>0", form1.adoc, adOpenDynamic, adLockReadOnly)
While Not rtmp.EOF
  If InStr(rtmp!Owner, "anredew") > 0 Then
    If pax$ = "x" Or pax$ = "w" Then Call anadd(trm(rtmp!wert))
  Else
    If InStr(rtmp!Owner, "anredem") > 0 Then
      If pax$ = "x" Or pax$ = "m" Then Call anadd(trm(rtmp!wert))
    Else
      Call anadd(trm(rtmp!wert))
    End If
  End If
  rtmp.MoveNext
Wend
End Sub
Private Sub anadd(l$)
Dim i%

'd2infile = "shwAdrDetail": d2insub = "anadd"
For i% = 0 To Anrede.ListCount - 1
  If Anrede.List(i%) = l$ Then Exit Sub
Next i%
Anrede.AddItem l$
End Sub

Private Sub abadd(l$)
Dim i%

'd2infile = "shwAdrDetail": d2insub = "anadd"
For i% = 0 To Abrede.ListCount - 1
  If Abrede.List(i%) = l$ Then Exit Sub
Next i%
Abrede.AddItem l$
End Sub

Private Sub anumsel_Click()
If knt_sav.Enabled Then Call savecheck
datf(6).text = anumsel.text
End Sub

Private Sub anumsel_DropDown()
Dim c$, rtmp As ADODB.Recordset, rrr, id$

anumsel.Clear
id$ = datf(0).text
If id$ = "" Then Exit Sub
c$ = "SELECT num FROM opt_allenummern where vid='" + id$ + "' and kid='-1' and numtyp='email'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
While Not rtmp.EOF
  If trm(datf(6).text) <> trm(rtmp!num) Then anumsel.AddItem trm(rtmp!num)
  rtmp.MoveNext
Wend
anumsel.AddItem datf(6).text
Call form1.chkallnums(id$, "-1", "email", datf(6).text)

End Sub

Private Sub Check1_Click()

'd2infile = "shwAdrDetail": d2insub = "Check1_Click"
Call savecheck
If Check1.value = 1 Then
  Command6.Visible = False
  knt_del.Enabled = False
Else
  Command6.Visible = True
  knt_del.Enabled = True
End If

End Sub

Private Sub Check2_Click()

'd2infile = "shwAdrDetail": d2insub = "Check2_Click"
Call savecheck
If Check2.value = 1 Then
'  Command13.Enabled = False
  Command13.Visible = False
  adr_del.Enabled = False
Else
'  Command13.Enabled = True
  Command13.Visible = True
  adr_del.Enabled = True
End If

End Sub

Private Sub Check4_Click()

DoEvents
If c4no Then Exit Sub
If Check4.value = 1 Then
  Call form1.setusersetting("zusatzinfos", "erweitert")
Else
  Call form1.setusersetting("zusatzinfos", "normal")
End If

End Sub

Private Sub Combo2_Click()
'd2infile = "shwAdrDetail": d2insub = "Combo2_Click"
Load tplan
tplan.setcaption (" - Projekt")
Call tplan.SetFocus
tplan.Text2.text = Combo2.List(Combo2.ListIndex)
DoEvents
If InStr(tplan.List6.List(1), transe("Neuer Auftritt")) = 0 Then
  Call tplan.Command26_Click
End If
If tplan.List6.ListCount > 1 Then
  tplan.List6.ListIndex = 1
  DoEvents
  Call tplan.List6_DblClick
End If

End Sub

Private Sub Combo3_Change()

'd2infile = "shwAdrDetail": d2insub = "Combo3_Change"
If srchit% = 0 Then Exit Sub
form1.Combo1.text = Combo3.text

End Sub

Private Sub Combo3_Click()
Dim t$, l%, i%

'd2infile = "shwAdrDetail": d2insub = "Combo3_Click"
t$ = Combo3.List(Combo3.ListIndex)
datf(0).text = t$
l% = Len(datf(0).text)
For i% = 0 To form1.List1.ListCount - 1
  If Left$(form1.List1.List(i%), l%) = t$ Then
    form1.List1.ListIndex = i%
    Call form1.List1_DblClick
    Exit Sub
  End If
Next i%
End Sub

Private Sub Combo3_GotFocus()
If knt_sav.Enabled Then Call savecheck
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
'd2infile = "shwAdrDetail": d2insub = "Combo3_KeyPress"
srchit% = 1
End Sub

Public Sub Command10_Click()
Dim id$, kid$

'd2infile = "shwAdrDetail": d2insub = "Command10_Click"
id$ = datf(0).text
'If id$ = "" Then Exit Sub
Load dochist2
kid$ = "-1"
If klist.ListIndex >= 0 Then kid$ = idxlist.List(klist.ListIndex)
Call dochist2.setkrit(id$, kid$)
On Error Resume Next
Call dochist2.SetFocus
On Error GoTo 0

End Sub

Public Sub Command11_Click()
Dim r As ADODB.Recordset, p%, n%, idn$, i%, pa$, tmp$, s$, s0$, isper As Boolean
Dim atp$, neuid$, cmd$, rrr, j%, ls$, neuname$, nn$, vn$, givenid$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command11_Click"

Call savecheck
isper = False
If idshow.Caption = "" And datf(1).text <> "" Then
    n% = linesof(datf(1).text) + 1
    idn$ = trm(lineof(1, datf(1).text))
    pa$ = ""
    For i% = 0 To postanredea.ListCount - 1
      If postanredea.List(i%) = idn$ Then
        pa$ = idn$
        idn$ = trm(lineof(2, datf(1).text))
        Exit For
      End If
    Next i%
    s0$ = idn$
'    tmp$ = word1(idn$)
'    idn$ = word2bis(idn$) + ", " + tmp$
    neuid$ = InputBox(transe("Neuer Sortiername:"), transe("Neue Adresse anlegen"), idn$)
    neuid = strrepl(trm(neuid), "/", "_")
    neuid = strrepl(neuid, "'", "´")
    neuid = strrepl(neuid, "&", "_")
    If neuid$ <> "" Then
      If InStr(neuid$, "(") > 0 Then
        MsgBox transe("Unerlaubtes Zeichen in der ID.")
        Call Command11_Click
        Exit Sub
      End If

      cmd$ = "select * from adresse where id='" & neuid$ & "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If r.EOF Then
        s$ = "insert into adresse (id,name) values('" & neuid$ & "','" & idn$ & "')"
        Call form1.sqlqry(s$)
        s$ = s0$
        i% = 2: If pa$ <> "" Then i% = 3
'        While i% + 2 <= n% And trm(lineof(i% + 1, datf(1).Text)) <> ""
'          If s$ <> "" Then s$ = s$ & vbCrLf
'          s$ = s$ & trm(lineof(i%, datf(1).Text))
'          i% = i% + 1
'        Wend
        If isdigit(Right$(trm(lineof(i%, datf(1).text)), 1)) = 0 Then
        While i% + 2 <= n% And isdigit(Right$(trm(lineof(i%, datf(1).text)), 1)) = 0
          If s$ <> "" Then s$ = s$ & vbCrLf
          s$ = s$ & trm(lineof(i%, datf(1).text))
          i% = i% + 1
        Wend
        End If
        s$ = "update adresse set name='" & s$ & "' where id='" & neuid$ & "'"
        Call form1.sqlqry(s$)
        If pa$ <> "" Then
          s$ = "update adresse set postanrede='" & pa$ & "' where id='" & neuid$ & "'"
          Call form1.sqlqry(s$)
        End If
        If i% + 2 <= n% Then
          s$ = trm(lineof(i%, datf(1).text))
          s$ = "update adresse set strasse='" & s$ & "' where id='" & neuid$ & "'"
          Call form1.sqlqry(s$)
          j% = i% + 1
          Do
            s$ = trm(lineof(j%, datf(1).text))
            j% = j% + 1
          Loop Until s$ <> "" Or j% > n%
          If Left(s$, 2) = "D " Then s$ = Mid$(s$, 3)
          Call form1.sqlqry("update adresse set ort='" & ohnePLZ(s$) & "' where id='" & neuid$ & "'")
          Call form1.sqlqry("update adresse set plz='" & trm(nurdiePLZ(s$)) & "' where id='" & neuid$ & "'")
        End If
        i% = i% + 2
        While i% <= n%
          s$ = trm(lineof(i%, datf(1).text))
          ls$ = LCase(s$)
          If (InStr(ls$, "tel") > 0 Or InStr(ls$, "fon") > 0) And InStr(ls$, "telefax") = 0 Then
            p% = InStr(s$, ":")
            If p% > 0 Then
              s$ = trm(Mid$(s$, p% + 1))
            Else
              s$ = onlynums(s$)
            End If
            s$ = "update adresse set tel='" & s$ & "' where id='" & neuid$ & "'"
            Call form1.sqlqry(s$)
          End If
          If InStr(ls$, "fax") > 0 Then
            p% = InStr(s$, ":")
            If p% > 0 Then
              s$ = trm(Mid$(s$, p% + 1))
            Else
              s$ = onlynums(s$)
            End If
            s$ = "update adresse set fax='" & s$ & "' where id='" & neuid$ & "'"
            Call form1.sqlqry(s$)
          End If
          If InStr(ls$, "mobil") > 0 Or InStr(ls$, "handy") > 0 Then
            p% = InStr(s$, ":")
            If p% > 0 Then
              s$ = trm(Mid$(s$, p% + 1))
            Else
              s$ = onlynums(s$)
            End If
            s$ = "update adresse set handy='" & s$ & "' where id='" & neuid$ & "'"
            Call form1.sqlqry(s$)
          End If
          If InStr(ls$, "@") > 0 Then
            p% = InStr(s$, ":")
            If p% > 0 Then
              s$ = trm(Mid$(s$, p% + 1))
            End If
            s$ = "update adresse set email='" & s$ & "' where id='" & neuid$ & "'"
            Call form1.sqlqry(s$)
          End If
          If InStr(ls$, "www.") > 0 Or InStr(ls$, "http:") > 0 Then
            p% = InStr(s$, ":")
            While p% > 0
              s$ = trm(Mid$(s$, p% + 1))
              p% = InStr(s$, ":")
            Wend
            While Left$(s$, 1) = "/": s$ = Mid$(s$, 2): Wend
            s$ = "update adresse set url='" & s$ & "' where id='" & neuid$ & "'"
            Call form1.sqlqry(s$)
          End If
          i% = i% + 1
        Wend
        form1.sqlqry _
          ( _
            "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + neuid$ + "','Person',NULL,'-1')" _
          )
      Else
        MsgBox transe("Dieser Sortiername existiert bereits.")
        Exit Sub
      End If
    End If
    datf(1).text = ""
    Me.BackColor = form1.cleancolor()
    knt_sav.Enabled = False
    adr_sav.Enabled = False
    If form1.getusersetting("internaldefault", "nein") = "ja" Then
      s$ = "update adresse set optinternal=1 where id='" & neuid$ & "'"
      Call form1.sqlqry(s$)
    End If
    Call refreshadrdetail(neuid$, "")
Else
  neuid$ = InputBox(transe("Neuer Sortiername:"), transe("Neue Adresse anlegen"), "")
  givenid$ = neuid$
  neuid$ = strrepl(trm(neuid$), "/", "_")
  neuid$ = strrepl(neuid$, "'", "`")
  neuid = strrepl(neuid, "&", "_")
  If neuid$ <> "" Then
    If InStr(neuid$, "(") > 0 Then
      MsgBox transe("Unerlaubtes Zeichen in der ID.")
      Call Command11_Click
      Exit Sub
    End If
    cmd$ = "select * from adresse where id='" + neuid$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If r.EOF Then
      neuname$ = givenid$
      p% = InStr(givenid$, ",")
      If p% > 0 Then
        nn$ = trm(Left$(givenid$, p% - 1))
        vn$ = trm(Mid$(givenid$, p% + 1))
        neuname$ = vn$ & " " & nn$
        isper = True
      End If
      atp$ = form1.grantadrtyp()
      If atp$ <> "" Then
        List1.AddItem transe(atp$)
      End If
      s$ = "insert into adresse (id,name) values('" + neuid$ + "','" + neuname$ + "')"
      form1.sqlqry (s$)
    End If
    If form1.getusersetting("internaldefault", "nein") = "ja" Then
      s$ = "update adresse set optinternal=1 where id='" & neuid$ & "'"
      Call form1.sqlqry(s$)
    End If
    Call refreshadrdetail(neuid$, "")
    If atp$ <> "" Then
      List1.AddItem transe(atp$)
      Call addtyp(atp$)
    End If
    If isper Then Call addtyp("Person")
  End If
End If

End Sub

Private Sub Command12_Click()
'd2infile = "shwAdrDetail": d2insub = "Command12_Click"
Call savecheck
klist.ListIndex = -1
currenk = "-1"
Command12.Enabled = False
Command36.Enabled = False
Label52.Caption = transe("inkl. Kontakte")
Unload zusinf
Unload bezlist
Call kposbuttonset
Call refreshadrdetail(datf(0).text, "-1")

End Sub

Private Sub Command13_Click()
Dim i%, up$, cmd$, id$, c$

'd2infile = "shwAdrDetail": d2insub = "Command13_Click"
Call savecheck

Command13.Visible = False
Check2.value = 1

id$ = datf(0).text
If id$ = "" Then Exit Sub
If klist.ListCount > 0 Then
  MsgBox transe("Löschen Sie zuerst alle Kontaktpersonen.")
  Exit Sub
End If
If List1.ListCount > 0 Then
  MsgBox transe("Löschen Sie zuerst alle Adresskategorien.")
  Exit Sub
End If

cmd$ = "delete from adresstyp where vid='" + id$ + "' and kid='-1';"
form1.sqlqry (cmd$)
cmd$ = "delete from adresse where id='" + id$ + "'"
form1.sqlqry (cmd$)
cmd$ = "delete from auftritthigru where auftrittsid='" + id$ + "'"
form1.sqlqry (cmd$)
cmd$ = "delete from anreden where kid='-1." + id$ + "'"
form1.sqlqry (cmd$)
If Not form1.isfieldmissing("opt_prios", "id") Then
  c$ = "delete from opt_prios where evnt='A:" + id$ + "'"
  Call form1.sqlqry(c$)
End If

Call nulldsp
End Sub

Public Sub Command14_Click()
Dim c$, wd, klrv%

'd2infile = "shwAdrDetail": d2insub = "Command14_Click"
gd1upd% = 0
DoEvents
Combo1.text = transe("keine Liste")
If Command14.Caption = transe("ohne Zusätze") Then
  Unload zusinf
  Unload bezlist
  c$ = transe("Zusatz-Infos")
  Command14.Caption = c$
  wd = wd0
  shwAdrDetail.Width = wd
  adr_zus.Checked = False
  gd1show.value = 0
Else
  klrv% = Val(form1.mylastFormVar(Me.name, "gd1show", "0"))
  If klrv% <> 0 Then klrv% = 1
  gd1show.value = klrv%
  DoEvents
  c$ = transe("ohne Zusätze")
  wd = wd1
  shwAdrDetail.Width = wd
  Command14.Caption = c$
  adr_zus.Checked = True
  Call rlist4
  Call rlist3
  Call rlist2
  Label36.Caption = form1.inmylanguage("Termine: ") + trm(List2.ListCount)
  DoEvents
  Call rcombo2
End If
gd1upd% = 1
Call mlist

End Sub

Private Sub Command15_Click()
Dim id$, i%, vorlage$, vtxt$, o%, l$, dtg$, vid$, kid$, wcu$, dtz$, l4$
Dim rtmp As ADODB.Recordset, rrr, c$, calid%, logid%, xld$

'd2infile = "shwAdrDetail": d2insub = "Command15_Click"
If Combo1.text = "CSV-Export" Or Combo1.text = "GEMA-Export" Then
  On Error Resume Next
  Kill form1.mydir() + "\_Termine.csv"
  rrr = Err
  On Error GoTo 0
  If rrr = 70 Then
    MsgBox "Die Datei " + form1.mydir() + "\_Termine.csv ist noch geöffnet und kann nicht erstellt werden."
    Exit Sub
  End If
  If Combo1.text = "GEMA-Export" Then
    xld$ = form1.getusersetting("exceldelimiter", ",")
    o% = FreeFile
    Open form1.mydir() + "\_Termine.csv" For Append As #o%
    Print #o%, """Datum""" + xld$ + """Art""" + xld$ + """Ort und Halle""";
    Print #o%, xld$ + """Veranstalter/Medium""";
    Print #o%, xld$ + """Programm mit GEMA-Nummer"""
    Close #o%
  End If
End If
form1.listenhauptperson = ""
form1.hordexlock = True
Call rlist2
form1.hordexlock = False
Label36.Caption = form1.inmylanguage("Termine: ") + trm(List2.ListCount)
DoEvents
Call rcombo2
If Combo1.text = "CSV-Export" Or Combo1.text = "GEMA-Export" Then
  X = Shell("explorer.exe " & form1.s0dir() & "\" + form1.docs() + "\" & form1.getuserid(), vbNormalFocus)
  Exit Sub
End If
If Combo1.text = "Cloud-Export" Then Combo1.text = transe("keine Liste")
If trm(Combo1.text) = "" Or LCase(transo(trm(Combo1.text))) = "keine liste" Or Combo1.text = "Cloud-Export" Then Exit Sub
form1.honorarlcount% = 0
Call form1.setAuftrittsdruckFuerAdresse(datf(0).text)
If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then
  Load dbupgrade
  dbupgrade.Caption = "Dokument wird erstellt ..."
  Call dbupgrade.SetFocus
End If
l4$ = datf(0).text
Call form1.auftrittsdruck(datf(0).text, form1.vorlagenverzeichnis() + "\adressen_" & Combo1.text & ".rtf", "adresse", l4$)
Call form1.setAuftrittsdruckFuerAdresse("")
Unload tplan

End Sub

Private Sub Command16_Click()
Dim i%, trgp$, eadr$

'd2infile = "shwAdrDetail": d2insub = "Command16_Click"
MousePointer = 11
trgp$ = ""
If usempth.value = 1 Then
  On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
  On Error GoTo 0
  trgp$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
End If
i% = klist.ListIndex
If i% < 0 Then
  eadr$ = "": If cadrpbez.text <> "Standard" Then eadr$ = "elseadr"
  Call form1.faxan(datf(0).text, "-1", form1.meinememovorlage(), "", "", trgp$, eadr$)
Else
  Call form1.faxan(datf(0).text, idxlist.List(klist.ListIndex), form1.meinememovorlage(), "", "", trgp$, "")
End If
MousePointer = 0
End Sub

Private Sub Command17_Click(Index As Integer)
Dim b$, l$, o%, eml$, knt$, rrr, sbj$

'd2infile = "shwAdrDetail": d2insub = "Command17_Click"
Call savecheck
If Index = 0 Then
  eml$ = trm(datf(6).text)
  knt$ = "-1"
Else
  eml$ = trm(kdat(5).text)
  knt$ = kdat(0).text
End If
If eml$ = "" Then Exit Sub
Load smtp
smtp.Visible = True
smtp.txtSendTo = eml$
smtp.adrid = datf(0).text
smtp.kid = knt$
Call smtp.txtMessageSubject.SetFocus
smtp.txtServer.Enabled = False
smtp.txtMailFrom.Enabled = False
If trm(Anrede.text) <> "" Then smtp.txtMessageText.text = trm(Anrede.text) & vbCrLf & vbCrLf & vbCrLf
If trm(Abrede.text) <> "" Then smtp.txtMessageText.text = smtp.txtMessageText.text & trm(Abrede.text) & vbCrLf
smtp.txtMessageText.text = smtp.txtMessageText.text & form1.uname$ & vbCrLf
Call form1.signaturinclude
If klist.ListIndex >= 0 Then smtp.kid = idxlist.List(klist.ListIndex)
If form1.getusersetting("MailclientSendet", "nein") = "ja" Then
  sbj$ = InputBox(transe("Betreff:"), transe("Beschreibung"), "", 100, 100)
  If Len(trm(sbj$)) = 0 Then
    Unload smtp
    Exit Sub
  End If
  smtp.eclient = 1
  smtp.txtMessageSubject = sbj$
  DoEvents
  Call smtp.cmdSend_Click
  DoEvents
  Unload smtp
End If
End Sub

Private Sub Command18_Click()
Dim c$, wd

'd2infile = "shwAdrDetail": d2insub = "Command18_Click"
If Command18.Caption = transe("ohne Medien") Then
  c$ = transe("Media Dateien")
  wd = hg0
  Me.Top = mytopmerk
  adr_med.Checked = False
Else
  c$ = transe("ohne Medien")
  wd = hg1
  mytopmerk = Me.Top
  If Me.Top + hg1 > Screen.Height Then
    Me.Top = Me.Top - (hg1 - hg0)
  End If
  adr_med.Checked = True
  Call mlist
End If

shwAdrDetail.Height = wd
Command18.Caption = c$
Call mlist

End Sub

Private Sub Command19_Click()
'd2infile = "shwAdrDetail": d2insub = "Command19_Click"
p1offs% = p1offs% + 1

Call mlist
End Sub

Public Sub Command2_Click()
Dim i%, trgp$, eadr$
'd2infile = "shwAdrDetail": d2insub = "Command2_Click"
Call savecheck
MousePointer = 11
i% = klist.ListIndex
trgp$ = ""
If usempth.value = 1 Then
  On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
  On Error GoTo 0
  trgp$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
End If
If i% < 0 Then
  eadr$ = "": If cadrpbez.text <> "Standard" Then eadr$ = "elseadr"
  Call form1.faxan(datf(0).text, "-1", form1.meinefaxvorlage(), "", "", trgp$, eadr$)
Else
  Call form1.faxan(datf(0).text, idxlist.List(klist.ListIndex), form1.meinefaxvorlage(), "", "", trgp$, "")
End If
MousePointer = 0
End Sub


Private Sub Command20_Click()
'd2infile = "shwAdrDetail": d2insub = "Command20_Click"
p1offs% = p1offs% - 1
If p1offs% < 0 Then p1offs% = 0
Call mlist
End Sub

Private Sub Command21_Click()
'd2infile = "shwAdrDetail": d2insub = "Command21_Click"
p1offs% = p1offs% - (p1cmdmax% + 1)
If p1offs% < 0 Then p1offs% = 0
Call mlist
End Sub

Private Sub Command22_Click()
'd2infile = "shwAdrDetail": d2insub = "Command22_Click"
p1offs% = p1offs% + p1cmdmax% + 1

Call mlist
End Sub

Private Sub Command1_Click()
'd2infile = "shwAdrDetail": d2insub = "Command1_Click"
Hide
On Error Resume Next
Unload vwr
Unload adrtypselector
Unload dochist2
Unload shwAdrDetail
Unload zusinf
Unload bezlist
Unload repertoire
On Error GoTo 0

End Sub

Private Sub Command23_Click()
Dim mpth$, ftm$, wert$, ask As Integer, l$, t$, ttest$, pb%
Dim bkmstart$, bkmend$, o%, p%, rrr, q%, rev$, ln$

'd2infile = "shwAdrDetail": d2insub = "Command23_Click"
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error GoTo 0
mpth$ = form1.s0dir() & "\" & form1.medien() & "\" & form1.medienname(datf(0).text) & "\" & form1.medienname(datf(0).text) & ".rtf"
If exist(form1.vorlagenverzeichnis() + "\media_bio.rtf") = 0 Then
  MsgBox transe("Vorlage unbekannt:") + " media_bio.rtf"
  Exit Sub
End If

wert$ = form1.saveasBox(mpth$)
wert$ = strrepl(wert$, " ", "_")
wert$ = strrepl(wert$, "/", "_")
If trm(wert$) = "" Then Exit Sub

If InStr(LCase(wert$), ".rtf") = 0 Then wert$ = wert$ + ".rtf"

If exist(mpth$ + "\" + wert$) = 1 Then
  ask = MsgBox(wert$ + ": " + transe("Die Datei existiert bereits. Überschreiben?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Vorhandene Datei löschen?"))
  If ask = vbNo Then Exit Sub
End If
o% = FreeFile
Open form1.vorlagenverzeichnis() + "\media_bio.rtf" For Input As #o%
p% = FreeFile
On Error Resume Next
Open wert$ For Output As #p%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox transe("Fehler #") & rrr & " " & Error$(rrr)
  Exit Sub
End If
While Not EOF(o%)
  Line Input #o%, l$
  While Len(l$) > 0
    q% = InStr(l$, bkmstart$)
    If q% > 0 Then
      t$ = Mid$(l$, q% + Len(bkmstart$))
      Print #p%, Left$(l$, q% - 1)
      t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
      rev$ = "": ttest$ = t$

      If InStr(ttest$, "__") > 0 Then
        rev$ = Mid$(ttest$, InStr(ttest$, "__") + 2)
        ttest$ = Left$(ttest$, InStr(ttest$, "__") - 1)
      End If
      If ttest$ = "user" Then
        ftm$ = form1.getusersetting(rev$)
        Print #p%, strrepl(ftm$, "\", "\\");
      End If
      If ttest$ = "system" Then
        Select Case LCase(rev$)
          Case "datum": Print #p%, Date
          Case "zeit": Print #p%, Left(Time, 5)
          Case "mwst": Print #p%, fixeurnozerotail(form1.sys_mwst / 100)
          Case Else: Print #p%, form1.getsystemsetting(rev$)
        End Select
      End If
      ln$ = Mid$(l$, q% + 1)
      Do
        pb% = InStr(LCase(ln$), bkmend$ + LCase(t$))
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
Wend
Close #p%
Close #o%
Call mlist
Call form1.openthisdoc(wert$, "")

End Sub

Private Sub Command24_Click()
Dim X

'd2infile = "shwAdrDetail": d2insub = "Command24_Click"
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
X = Shell("explorer.exe " + form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text), vbNormalFocus)
On Error GoTo 0

End Sub

Private Sub Command25_Click()

'd2infile = "shwAdrDetail": d2insub = "Command25_Click"
Call savecheck
Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), datf(0).text & " [Wiedervorlage] Adresse:" + _
               datf(0).text, "", Date, Left(Time, 5))
Call create2do.SetFocus
create2do.Text1(1).Enabled = False
create2do.Text1(3).Enabled = False
End Sub

Private Sub Command26_Click()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Command26_Click"
  id$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 id$ + "','-1'" + _
                 ",'Neuer Auftritt','" + transe("Neuer Auftritt") + "','" + _
                 datum2sql(CDate(Date)) + "')")
  Unload auftritt
  DoEvents
  Load auftritt
  Call auftritt.SetFocus
  Call auftritt.showrec(id$, 0)

End Sub

Private Sub Command27_Click()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "Command27_Click"
id$ = datf(0).text
Load bplan
On Error Resume Next
Call bplan.SetFocus
On Error GoTo 0
bplan.hid.text = id$
End Sub

Private Sub Command28_Click()

'd2infile = "shwAdrDetail": d2insub = "Command28_Click"
Call form1.handbuchcall("06-Adressen.htm")

End Sub

Private Sub Command29_Click()
'd2infile = "shwAdrDetail": d2insub = "Command29_Click"
Call Command24_Click
End Sub

Private Sub Command3_Click()
Dim i%, trgp$, vorl$, eadr$

'd2infile = "shwAdrDetail": d2insub = "Command3_Click"
Call savecheck
i% = klist.ListIndex
MousePointer = 11
trgp$ = ""
If usempth.value = 1 Then
  On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
  On Error GoTo 0
  trgp$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
End If
vorl$ = altbvorl.text
If vorl$ = "" Then vorl$ = form1.meinebriefvorlage()
If i% < 0 Then
  eadr$ = "": If cadrpbez.text <> "Standard" Then eadr$ = "elseadr"
  Call form1.faxan(datf(0).text, "-1", vorl$, "", "", trgp$, eadr$)
Else
  eadr$ = "": If ckadrpbez.text <> "Standard" Then eadr$ = "elsekadr"
  Call form1.faxan(datf(0).text, idxlist.List(klist.ListIndex), vorl$, "", "", trgp$, eadr$)
End If
MousePointer = 0

End Sub

Private Sub Command30_Click()
'd2infile = "shwAdrDetail": d2insub = "Command30_Click"
Load splan
On Error Resume Next
splan.hid.text = datf(0).text
Call splan.SetFocus
On Error GoTo 0

End Sub

Private Sub Command31_Click()
Dim tpid$, c$, tb$, o%, tg$

'd2infile = "shwAdrDetail": d2insub = "Command31_Click"
tpid$ = datf(0).text
If tpid$ = "" Then Exit Sub

MousePointer = 11: DoEvents
On Error Resume Next
Kill form1.mydatadir() & "\*.sql"
On Error GoTo 0
Call form1.sqlex_adresse("adresse", "id", tpid$)
Load smtp
On Error Resume Next
Call smtp.SetFocus
On Error GoTo 0
smtp.txtMessageSubject = "Agencyprof Datenpakete Adresse " & datf(0).text
smtp.txtMessageText = "Speichern Sie das Attachment in Ihrem Agencyprof-Verzeichnis"
tg$ = Dir(form1.mydatadir() & "\*.sql")
While tg$ <> ""
    Call smtp.attachfile(form1.mydatadir() & "\" & tg$)
    tg$ = Dir
Wend
MousePointer = 0

End Sub

Private Sub Command32_Click()
Dim mpth$, ftm$, id$, r As ADODB.Recordset, wert$, o%, t$, ttest$, ln$
Dim bkmstart$, bkmend$, ask As Integer, p%, rrr, l$, q%, rev$, pb%, c$, trma As String

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command32_Click"
id$ = trm(datf(0).text)
If id$ = "" Then Exit Sub
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error GoTo 0
mpth$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
If exist(form1.vorlagenverzeichnis() + "\media_bio.rtf") = 0 Then
  MsgBox transe("Vorlage unbekannt:") + " media_bio.rtf"
  Exit Sub
End If

wert$ = form1.saveasBox(mpth$ & "\" & datf(0).text)
wert$ = strrepl(wert$, " ", "_")
If trm(wert$) = "" Then Exit Sub

If Right$(LCase(wert$), 4) <> ".rtf" Then wert$ = wert$ + ".rtf"

If exist(wert$) = 1 Then
  ask = MsgBox(wert$ + ": " + transe("Die Datei existiert bereits. Überschreiben?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Vorhandene Datei löschen?"))
  If ask = vbNo Then Exit Sub
End If
o% = FreeFile
Open form1.vorlagenverzeichnis() + "\media_bio.rtf" For Input As #o%
p% = FreeFile
On Error Resume Next
Open wert$ For Output As #p%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Fehler #" & rrr & " " & Error$(rrr)
  Exit Sub
End If
While Not EOF(o%)
  Line Input #o%, l$
  While Len(l$) > 0
    q% = InStr(l$, bkmstart$)
    If q% > 0 Then
      t$ = Mid$(l$, q% + Len(bkmstart$))
      Print #p%, Left$(l$, q% - 1)
      t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
      rev$ = "": ttest$ = t$

      If InStr(ttest$, "__") > 0 Then
        rev$ = Mid$(ttest$, InStr(ttest$, "__") + 2)
        ttest$ = Left$(ttest$, InStr(ttest$, "__") - 1)
      End If
      If ttest$ = "user" Then
        ftm$ = form1.getusersetting(rev$)
        Print #p%, strrepl(ftm$, "\", "\\");
      End If
      If ttest$ = "system" Then
        Select Case LCase(rev$)
          Case "datum": Print #p%, Date
          Case "zeit": Print #p%, Left(Time, 5)
          Case "mwst": Print #p%, fixeurnozerotail(form1.sys_mwst / 100)
          Case Else: Print #p%, form1.getsystemsetting(rev$)
        End Select
      End If
      ln$ = Mid$(l$, q% + 1)
      Do
        pb% = InStr(LCase(ln$), bkmend$ + LCase(t$))
        If pb% = 0 Then Line Input #o%, ln$
      Loop Until pb% > 0
      ln$ = Mid$(ln$, pb%)
      If InStr(ln$, "}") = 0 Then
        l$ = ""
      Else
        l$ = Mid$(ln$, InStr(ln$, "}") + 1)
      End If
    Else
      If l$ = "\par }}" Then
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM adresse where id ='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          trma = trm(r!postanrede): If trma <> "" Then trma = trma + " "
          Print #p%, trma + trm(r!name) & "\par "
          Print #p%, trm(r!strasse) & "\par "
          If Not IsNull(r!plz) Then Print #p%, trm(r!plz) & " "
          Print #p%, trm(r!ort) & "\par "
          Print #p%, trm("Tel: " & r!tel) & "\par "
          Print #p%, trm("Handy: " & r!handy) & "\par "
          Print #p%, trm("Fax: " & r!fax) & "\par "
          Print #p%, trm("Mail: " & r!email) & "\par Kontakte:\par "
          c$ = "SELECT * FROM kontakt where vid ='" + id$ + "'"
          Set r = New ADODB.Recordset
          r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not r.EOF
            trma = trm(r!postanrede): If trma <> "" Then trma = trma + " "
            Print #p%, trma + trm(r!name) & ",  "
            Print #p%, trm("Tel: " & r!tel) & ", "
            Print #p%, trm("Handy: " & r!handy) & ", "
            Print #p%, trm("Fax  : " & r!fax) & ", "
            Print #p%, trm("Mail : " & r!email) & "\par "
            r.MoveNext
          Wend
        End If
      End If
      Print #p%, l$
      l$ = ""
    End If
  Wend
Wend
Close #p%
Close #o%
Call mlist
Call form1.openthisdoc(wert$, "")


End Sub

Private Sub Command33_Click()
Dim plz$, rest$, r As ADODB.Recordset, c$, rrr, i%, j%, plzk$, plzf$, lkz$
Dim konly As Boolean, clpt$, plzf1$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command33_Click"
Clipboard.Clear
DoEvents
rest$ = ""
plz$ = trm(datf(14).text & " " & datf(13).text)
konly = False
If klist.ListIndex >= 0 Then
  kid$ = idxlist.List(klist.ListIndex)
  konly = True
End If
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
If konly Then
  c$ = "SELECT * FROM kontakt where id='" + kid$ + "'"
Else
  c$ = "SELECT * FROM kontakt where vid='" + datf(0).text + "' order by name"
End If
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If rest$ <> "" Then rest$ = rest$ & vbCrLf
  For i% = 2 To 8
    If Not IsNull(r.Fields(i%).value) Then
      If i% > 2 And rest$ <> "" And Not konly Then If i% <> 7 Then rest$ = rest$ & ", "
      If konly Then
        If i% <> 7 Then
          If i% > 2 Then rest$ = rest$ & trm(r.Fields(i%).name) & ": "
          rest$ = rest$ & trm(r.Fields(i%).value) + vbCrLf
          If i% = 2 Then
            lkz$ = trm(r.Fields(11).value): If lkz$ = "" Then lkz$ = datf(2).text
            rest$ = rest$ & lkz$ + vbCrLf
            lkz$ = trm(r.Fields(12).value)
            plzk$ = trm(lkz$ + " " + trm(trm(r.Fields(13).value) + " " + trm(r.Fields(14).value)))
            plzf1$ = trm(trm(r.Fields(15).value) + " " + trm(r.Fields(16).value))
            plzf$ = "": If plzf1$ <> "" Then plzf$ = trm(lkz$ + " " + plzf1$)
            If plzk$ = "" And plzf$ = "" Then plzk$ = trm(plz$ & " " & datf(3).text)
            If plzk$ <> "" Then rest$ = rest$ + plzk$ + vbCrLf
            If plzf$ <> "" Then rest$ = rest$ + plzf$ + vbCrLf
          End If
        End If
      Else
        If i% <> 7 Then rest$ = rest$ & trm(r.Fields(i%).name & ": " & r.Fields(i%).value)
      End If
    End If
  Next i%
  r.MoveNext
Wend
If konly Then
  clpt$ = datf(1).text & vbCrLf
Else
  plzf1$ = trm(trm(plzp.text) & " " & trm(postf.text))
  If plzf1$ <> "" Then plzf1$ = trm(trm(datf(14).text) & " " & plzf1$) & vbCrLf
  clpt$ = datf(1).text & vbCrLf & _
         datf(2).text & vbCrLf & _
         trm(plz$ & " " & datf(3).text) & vbCrLf & plzf1$ & _
         "Tel.: " & datf(4).text & vbCrLf & _
         "Fax: " & datf(5).text & vbCrLf & _
         "Email: " & datf(6).text & vbCrLf & _
         "Handy: " & datf(9).text & vbCrLf & _
         "URL: " & datf(10).text & vbCrLf
End If
clpt$ = clpt$ + rest$
Clipboard.Clear
Clipboard.settext clpt$
End Sub

Private Sub Command34_Click()
Dim fn$, id$, plz$, FullName$, em$, eml$, ohnehandy As Boolean, mfn$, o%
Dim fax$, tel$, cell$, url$, i%, org$, st$, ort$, Out$

'd2infile = "shwAdrDetail": d2insub = "Command34_Click"
id$ = datf(0).text
If id$ = "" Then Exit Sub
Call savecheck
MousePointer = 11: DoEvents
ohnehandy = False
If form1.getusersetting("vcfohnehandy") = "ja" Then ohnehandy = True
mfn$ = form1.mkfn(id$)
plz$ = trm(datf(14).text & " " & datf(13).text)

fn$ = form1.mydatadir() & "\" & mfn$ & ".vcf"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, "begin: vcard"
eml$ = trm(datf(6).text)
fax$ = trm(datf(5).text)
tel$ = trm(datf(4).text)
cell$ = trm(datf(9).text)
url$ = trm(datf(10).text)
i% = klist.ListIndex
If i% >= 0 Then
  FullName$ = kdat(2).text
  em$ = trm(kdat(5).text): If em$ <> "" Then eml$ = em$
  em$ = trm(kdat(3).text): If em$ <> "" Then tel$ = em$
  em$ = trm(kdat(4).text): If em$ <> "" Then fax$ = em$
  em$ = trm(kdat(6).text): If em$ <> "" Then cell$ = em$
  em$ = trm(kdat(8).text): If em$ <> "" Then url$ = em$
Else
  FullName$ = datf(1).text
End If
FullName$ = strrepl(FullName$, vbCrLf, " ")
Print #o%, "fn: " & FullName$
Print #o%, "n: " & FullName$
org$ = datf(1).text
org$ = strrepl(org$, vbCrLf, " ")
Print #o%, "org:" & org$
st$ = datf(2).text: st$ = strrepl(st$, vbCrLf, " ")
ort$ = datf(3).text: ort$ = strrepl(ort$, vbCrLf, " ")
Out$ = st$ & ";" & ort$ & ";;" & plz$: If trm(Out$) <> "" Then Print #o%, "adr;dom:;;" & Out$
Out$ = eml$: If trm(Out$) <> "" Then Print #o%, "email;internet:" & eml$
Out$ = tel$: If trm(Out$) <> "" Then Print #o%, "tel;work:" & tel$
Out$ = fax$: If trm(Out$) <> "" Then Print #o%, "tel;fax:" & fax$
If Not ohnehandy Then
  Out$ = cell$: If trm(Out$) <> "" Then Print #o%, "tel;cell:" & cell$
End If
Out$ = url$: If trm(Out$) <> "" Then Print #o%, "url:" & url$
Print #o%, "version:2.1"
Print #o%, "End: vcard"
Close #o%
Load smtp
On Error Resume Next
Call smtp.SetFocus
On Error GoTo 0
smtp.txtMessageSubject = "Adresse von " & datf(0).text & "(vCard aus Agencyprof)"
smtp.txtMessageText = "vCard aus Agencyprof (http://www.agencyprof.de)."
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Die angehängte vCard enthält die folgenden Daten:" & vbCrLf
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Voller Name: " & FullName$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Name: " & FullName$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Firma: " & org$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Adresse: " & st$ & ";" & ort$ & ";;" & plz$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Email;internet:" & eml$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Telefon:" & tel$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Fax:" & fax$
If Not ohnehandy Then smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "Handy:" & cell$
smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "URL:" & url$ & vbCrLf
Call smtp.attachfile(fn$)
MousePointer = 0

End Sub

Private Sub Command35_Click()
Dim rtmp As ADODB.Recordset, tel$, hnd$, i%, l$, le$, c$
Dim s As ADODB.Recordset, rrr, na$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command35_Click"
MousePointer = 11: DoEvents

Load dialselect
Call dialselect.SetFocus
dialselect.List1.Clear
If trm(datf(4).text) <> "" Then dialselect.List1.AddItem transe("Tel wählt:") + " " & datf(4).text
If trm(datf(9).text) <> "" Then dialselect.List1.AddItem transe("Handy wählt:") + " " & datf(9).text
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
c$ = "SELECT * FROM auftritthigru where auftrittsid ='" + idshow.Caption & "' and instr(feldname,'Tel-')>0"
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not s.EOF
  dialselect.List1.AddItem s!feldname & " " + transe("wählt:") + " " & s!felddaten
  s.MoveNext
Wend
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,name,tel,handy FROM kontakt where vid ='" + idshow.Caption & "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  rtmp.MoveFirst
  While Not rtmp.EOF
    na$ = trm(rtmp!name)
    tel$ = trm(rtmp!tel)
    hnd$ = trm(rtmp!handy)
    If tel$ <> "" Then dialselect.List1.AddItem transe("Kontakt Tel   ") & na$ & " " + transe("wählt:") + " " & tel$
    If hnd$ <> "" Then dialselect.List1.AddItem transe("Kontakt Handy ") & na$ & " " + transe("wählt:") + " " & hnd$
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
    c$ = "SELECT * FROM auftritthigru where auftrittsid ='" + idshow.Caption & rtmp!id & "' and instr(feldname,'Tel-')>0"
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not s.EOF
      dialselect.List1.AddItem na$ & " " & s!feldname + " " + transe("wählt:") + " " & s!felddaten
      s.MoveNext
    Wend
    rtmp.MoveNext
  Wend
End If
If Not form1.isfieldmissing("opt_numbers", "id") Then
  dialselect.List1.AddItem " ---- "
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  c$ = "SELECT id,kid,NumTyp,Number from opt_numbers where vid ='" + idshow.Caption & "'"
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    rtmp.MoveFirst
    While Not rtmp.EOF
      na$ = trm(rtmp!NumTyp)
      If trm(rtmp!kid) <> "-1" Then
        na$ = form1.get_kontaktname_by_id(trm(rtmp!kid)) + " " + na$
      End If
      tel$ = trm(rtmp!Number)
      If tel$ <> "" Then dialselect.List1.AddItem strrepl(na$, ":", "") & " " + transe("wählt:") + " " & tel$ + Space$(100) + ":" + rtmp!id
      rtmp.MoveNext
    Wend
  End If
End If
hnd$ = datf(7).text
le$ = lastlineof(hnd$)
i% = 1: tel$ = "$/(hoffentlichkommtdasnievor+$%("
Do
  l$ = lineof(i%, hnd$)
  If InStr(l$, "Tel:") > 0 Or (InStr(l$, "Tel-") > 0 And InStr(l$, ":") > 0) Then
    dialselect.List1.AddItem transe("aus Hinweis wählt:") + " " & l$
  End If
  i% = i% + 1
Loop Until l$ = le$
tel$ = lineof(i%, hnd$)

MousePointer = 0
End Sub

Private Sub Command36_Click()
Dim V%

'd2infile = "shwAdrDetail": d2insub = "Command36_Click"
Call savecheck
Command36.Enabled = False
V% = 0: If trm(kadat(V%).text) = "" Then kadat(V%).text = datf(2).text
V% = 1: If trm(kadat(V%).text) = "" Then kadat(V%).text = datf(14).text
V% = 2: If trm(kadat(V%).text) = "" Then kadat(V%).text = datf(13).text
V% = 4: If trm(kadat(V%).text) = "" Then kadat(V%).text = plzp.text
V% = 3: If trm(kadat(V%).text) = "" Then kadat(V%).text = datf(3).text
V% = 5: If trm(kadat(V%).text) = "" Then kadat(V%).text = postf.text
V% = 3: If trm(kdat(V%).text) = "" Then kdat(V%).text = datf(4).text
V% = 4: If trm(kdat(V%).text) = "" Then kdat(V%).text = datf(5).text
V% = 5: If trm(kdat(V%).text) = "" Then kdat(V%).text = datf(6).text
V% = 6: If trm(kdat(V%).text) = "" Then kdat(V%).text = datf(9).text
V% = 8: If trm(kdat(V%).text) = "" Then kdat(V%).text = datf(10).text

End Sub

Private Sub Command37_Click()
Dim tpid$, c$, tb$, o%, tg$, neuid$

'd2infile = "shwAdrDetail": d2insub = "Command37_Click"
Call savecheck
tpid$ = datf(0).text
If tpid$ = "" Then Exit Sub
neuid$ = trm(InputBox(transe("Neuer Sortiername:"), transe("Neue Adresse anlegen"), ""))
If neuid$ = "" Then Exit Sub
MousePointer = 11: DoEvents
On Error Resume Next
Kill form1.mydatadir() & "\*.sql"
On Error GoTo 0
Call form1.sqlex_adresseas("adresse", "id", tpid$, neuid$)
tg$ = Dir(form1.mydatadir() & "\*.sql")
While tg$ <> ""
  Call FileCopy(form1.mydatadir() & "\" & tg$, form1.s0dir() & "\" & tg$)
  Kill form1.mydatadir() & "\" & tg$
  On Error GoTo 0
  tg$ = Dir
Wend
Load agx
Call agx.Command3_Click: DoEvents
Call agx.Command2_Click: DoEvents
Unload agx
form1.Combo1.text = neuid$
MousePointer = 0

End Sub

Private Sub Command38_Click()

'd2infile = "shwAdrDetail": d2insub = "Command38_Click"
Load landwahl
landwahl.Show
Call landwahl.settarget(1)
On Error Resume Next
Call landwahl.SetFocus
On Error GoTo 0

End Sub

Private Sub Command39_Click()
'd2infile = "shwAdrDetail": d2insub = "Command39_Click"
Load landwahl
landwahl.Show
Call landwahl.settarget(2)
On Error Resume Next
Call landwahl.SetFocus
On Error GoTo 0

End Sub

Public Sub Command4_Click()
Dim i%, u1$, u2$, cmd$, rtmp As QueryDef, up$, fn$, restoredata$, ww$, antw
Dim s As ADODB.Recordset, r As ADODB.Recordset, id$, c$, rrr, ask As Integer, anid$
Dim iu As Integer, fldnams$, Index As Integer, tmpb As Boolean

Call form1.dbg2f("Command4_Click", "shwadrdetail", "Command4_Click")
Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command4_Click"
id$ = datf(0).text
If id$ = "" Then Exit Sub
'länder immer gross?
If form1.getusersetting("laenderimmergross", "nein") = "ja" Then
  Index = 14
  If datf(Index).text <> UCase(datf(Index).text) Then
    datf(Index).text = UCase(datf(Index).text)
    DoEvents
  End If
End If
Call form1.chkallnums(id$, "-1", "email", datf(6).text)
If cadrpbez.text <> "Standard" Then
  up$ = trm(InputBox(transe("Sie speichern nicht die Standardadresse." + vbCrLf + "Die jetzt eingetragene Adresse wird die Standardadresse." + vbCrLf + "Bestätigen Sie bitte mit JA"), transe("Adresse unklar.")))
  If LCase(trm(up$)) <> "ja" Then Exit Sub
  Call svadr("Standard")
End If
MousePointer = 11
If form1.getusersetting("adressenstandautomatisch", "ja") = "ja" Then
  datf(8).text = trm(Date) + " " + trm(Time)
End If
up$ = form1.getusersetting("adresshistorie", "nein")
If up$ = "ja" Or up$ = "erweitert" Then
  'If up$ = "erweitert" Then
    restoredata$ = "N/A"
    fn$ = form1.s0dir() + "\tmp\" + form1.medienname(datf(0).text) + "." + id$
    i% = FreeFile
    'Open fn$ For Output As #i%: Close #i%
    c$ = "delete from dochist where adresse='" + id$ + "' and kontakt='-1' and doctyp='" + transe("Datenänderung") + "';"
    Call form1.sqlqry(c$)
    c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
            form1.newid("dochist", "id", 19) & "','" & id$ & "','-1','" + fn$ & "','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','" + transe("geändert") & "','" + transe("Datenänderung") + "')"
    Call form1.sqlqry(c$)
  'End If
End If
fldnams$ = "id,name,strasse,ort,tel,fax,email,hinweise,stand,handy,url,kdnr,telfaxhandy,plz,land,bundesland"
For i% = 0 To nflds
  u2$ = cut_d1(fldnams$, ","): fldnams$ = cut_d2bis(fldnams$, ",")
  If i% > 0 Then
    If i% = 14 Then
test2land:
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
      c$ = "SELECT * FROM sysvars where owner='sysvar_system_landeskennung_" & datf(14).text & "'"
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not s.EOF Then
        If datf(i%).text <> s!wert Then datf(14).text = s!wert
      Else
        ask = MsgBox(datf(i%).text & ": " + transe("Das Land ist bisher nicht bekannt.") & vbCrLf & transe("Ist es richtig geschieben?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Land aufnehmen?"))
        If ask = vbYes Then
          c$ = "insert into sysvars (id,owner,wert) values('" & form1.newid("sysvars", "id", 22) & "','sysvar_system_landeskennung_" & datf(i%).text & "','" & datf(i%).text & "')": Call form1.sqlqry(c$)
        Else
          datf(i%).text = InputBox(transe("Geben Sie das Land ein:"), transe("Land?"), "")
          If datf(i%).text = "" Then datf(i%).text = form1.getusersetting("MeinLand", "D")
          GoTo test2land
        End If
      End If
    End If
    cmd$ = "update adresse set " + u2$ + " = '"
    If i% = 8 Then
      c$ = datum2sql(word1(datf(i%).text)) + " " + word2bis(datf(i%).text)
      If c$ = " " Then c$ = datum2sql(Date) + " " + trm(Time)
      cmd$ = cmd$ + c$
    Else
      cmd$ = cmd$ + strrepl(datf(i%).text, "'", "´")
    End If
    cmd$ = cmd$ + "' where id='" + id$ + "'"
    Call form1.sqlqry(cmd$)
  End If
Next i%
u2$ = Left$(u2$, Len(u2$) - 1)
iu = intuse.value

'If trm(plzp.text) <> "" Then
  cmd$ = "update adresse set plzpostfach='" & plzp.text & "' where id='" & id$ & "'"
  Call form1.sqlqry(cmd$)
'End If
'If trm(postf.text) <> "" Then
  cmd$ = "update adresse set postfach='" & postf.text & "' where id='" & id$ & "'"
  Call form1.sqlqry(cmd$)
'End If
If trm(postanredea.text) <> "" Then
  cmd$ = "update adresse set postanrede='" & postanredea.text & "' where id='" & id$ & "'"
  Call form1.sqlqry(cmd$)
End If
If Not form1.isfieldmissing("addresse", "opttel") Then
  If opttel.text <> "" Then
    cmd$ = "update adresse set opttel='" & opttel.text & "' where id='" & id$ & "'"
    Call form1.sqlqry(cmd$)
  End If
End If
If klist.ListIndex < 0 Then
  form1.sqlqry ("delete from anreden where kid='-1." & id$ & "' and user='" & form1.anredeuser$ & "'")
  If Len(Anrede.text) > 0 Or Len(Abrede.text) > 0 Then
    form1.sqlqry ("insert into anreden (id,kid,user,an,ab) values('" & form1.newid("anreden", "id", 20) + "','-1." & id$ & "','" & form1.anredeuser$ & "','" & Anrede.text & " ','" & Abrede.text & " ')")
  End If
  If form1.getusersetting("systemanredensetzen", "nein") = "ja" Then
    form1.sqlqry ("delete from anreden where kid='-1." & id$ & "' and user='system'")
    anid$ = form1.newid("anreden", "id", 18)
    c$ = "insert into anreden (id,kid,user) values('" + anid$ + "','-1." + id$ + "','system')"
    Call form1.sqlqry(c$)
    If Len(Anrede.text) > 0 Then
      form1.sqlqry ("update anreden set an='" & Anrede.text & "' where id='" + anid$ + "'")
    End If
    If Len(Abrede.text) > 0 Then
      form1.sqlqry ("update anreden set ab='" & Abrede.text & "' where id='" + anid$ + "'")
    End If
  End If
End If
'wohnt bei
c$ = "select * from adresstyp where vid='" + id$ + "' and typ='rel:wohnt bei'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
ww$ = ""
If Not s.EOF Then
  ww$ = s!wert
  c$ = "select * from adresse where id='" + ww$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    If trm(r!strasse) <> trm(datf(2).text) Or _
       trm(r!ort) <> trm(datf(3).text) Or _
       trm(datf(13).text) <> trm(r!plz) Or _
       trm(plzp.text) <> trm(r!plzpostfach) Or _
       trm(postf.text) <> trm(r!postfach) Then
      antw = MsgBox(id$ + " " + transe("wohnt bei") & " " & ww$ & vbCrLf & transe("Ist dies ein Auszug?") + vbCrLf + _
                    transe("[Ja] entfernt die wohnt-bei-Beziehung von") + " " + id$ + vbCrLf + _
                    transe("[Nein] ändert die Adresse von") + " " + ww$, vbYesNo + vbCritical + vbDefaultButton2, transe("Sortiernamen ändern?"))
      If antw = vbYes Then
        c$ = "delete from adresstyp where vid='" + id$ + "' and typ='rel:wohnt bei'"
        Call form1.sqlqry(c$)
      Else
        c$ = "update adresse set strasse='" + datf(2).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set ort='" + datf(3).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set plz='" + datf(13).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set plzpostfach='" + plzp.text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set postfach='" + postf.text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
      End If
    End If
  End If
End If
c$ = "select * from adresstyp where wert='" + id$ + "' and typ='rel:wohnt bei'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
ww$ = ""
While Not s.EOF
  ww$ = s!vid
  c$ = "select * from adresse where id='" + ww$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    If trm(r!strasse) <> trm(datf(2).text) Or _
       trm(r!ort) <> trm(datf(3).text) Or _
       trm(datf(13).text) <> trm(r!plz) Or _
       trm(plzp.text) <> trm(r!plzpostfach) Or _
       trm(postf.text) <> trm(r!postfach) Then
      antw = MsgBox(ww$ + " " + transe("wohnt bei") & " " & id$ & vbCrLf & transe("Ist dies ein Auszug?") + vbCrLf + _
                    transe("[Ja] entfernt die wohnt-bei-Beziehung von") + " " + ww$ + vbCrLf + _
                    transe("[Nein] ändert die Adresse von") + " " + ww$, vbYesNo + vbCritical + vbDefaultButton2, transe("Sortiernamen ändern?"))
      If antw = vbYes Then
        c$ = "delete from adresstyp where vid='" + ww$ + "' and typ='rel:wohnt bei'"
        Call form1.sqlqry(c$)
      Else
        c$ = "update adresse set strasse='" + datf(2).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set ort='" + datf(3).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set plz='" + datf(13).text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set plzpostfach='" + plzp.text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set postfach='" + postf.text + "' where id='" + ww$ + "'": Call form1.sqlqry(c$)
      End If
    End If
  End If
  s.MoveNext
Wend

Call refreshadrdetail(id$, "")
If form1.cloud Then
  ww$ = form1.cloudmanager + form1.cloudstaff: c$ = "x"
  While ww$ <> ""
    c$ = cut_d1(ww$, "|"): ww$ = cut_d2bis(ww$, "|")
    If c$ <> "" Then
      tmpb = form1.cloudcreateadr(id$, "", c$)
    End If
  Wend
  Call form1.adr2cloud(id$)
End If
intuse.value = iu
MousePointer = 0
End Sub

Private Sub Command41_Click()
Dim V%

'd2infile = "shwAdrDetail": d2insub = "Command41_Click"
Call savecheck
Command36.Enabled = False
V% = 0: If trm(datf(2).text) = "" Then datf(2).text = kadat(V%).text
V% = 1: If trm(datf(14).text) = "" Then datf(14).text = kadat(V%).text
V% = 2: If trm(datf(13).text) = "" Then datf(13).text = kadat(V%).text
V% = 4: If trm(plzp.text) = "" Then plzp.text = kadat(V%).text
V% = 3: If trm(datf(3).text) = "" Then datf(3).text = kadat(V%).text
V% = 5: If trm(postf.text) = "" Then postf.text = kadat(V%).text
V% = 3: If trm(datf(4).text) = "" Then datf(4).text = kdat(V%).text
V% = 4: If trm(datf(5).text) = "" Then datf(5).text = kdat(V%).text
V% = 5: If trm(datf(6).text) = "" Then datf(6).text = kdat(V%).text
V% = 6: If trm(datf(9).text) = "" Then datf(9).text = kdat(V%).text
V% = 8: If trm(datf(10).text) = "" Then datf(10).text = kdat(V%).text

End Sub

Private Sub Command42_Click()
Dim aid$, kid$, kidx$

'd2infile = "shwAdrDetail": d2insub = "Command42_Click"
List4.Clear
If List1b.Visible Then
  List1b.Visible = False
  Kategorie.Caption = transe("Kategorien:")
  Unload zusinf
  Unload bezlist
Else
  List1b.Visible = True
  Kategorie.Caption = transe("Beziehungen:")
  Unload zusinf
  Unload bezlist
  DoEvents
  On Error Resume Next
  Load bezlist
  Call bezlist.SetFocus
  On Error GoTo 0
  kid$ = "-1": kidx$ = "-1"
  If klist.ListIndex >= 0 Then
    kid$ = klist.List(klist.ListIndex)
    kid$ = idxlist.List(klist.ListIndex)
  End If
  DoEvents
'  Call bezlist.setcurrent(idshow.Caption, kid$, kidx$)
  If List1b.ListCount > 0 Then List1b.ListIndex = 0
End If
End Sub

Private Sub Command43_Click()
Dim i%, up$, cmd$, id$, kid$, stmp As ADODB.Recordset, neukwert, neuwert, neuawert, c$
Dim rrr
Dim rtmp As ADODB.Recordset, sida$, sidk$, knam$, knid$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command43_Click"
id$ = datf(0).text
If id$ = "" Then Exit Sub
Call savecheck
Load adrselect
Call adrselect.sel_init(id$, "")
Call adrselect.SetFocus
Do
  DoEvents
Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
If adrselect.sel_brk() = 0 Then
  MousePointer = 11: DoEvents
  neukwert = adrselect.get_kontsel()
  neuwert = trm(adrselect.sel_getselected()): neuawert = neuwert
  Unload adrselect
  If neukwert <> "" Then
    'aus kontakt
    c$ = "update adresstyp set vid='" + id$ + "' where vid='" + neuwert + "' and kid='" + adrselect.kontselid + "'"
    Call form1.sqlqry(c$)
    c$ = "update auftritthigru set auftrittsid='" + id$ + adrselect.kontselid + "' where auftrittsid='" + neuwert + adrselect.kontselid + "'"
    Call form1.sqlqry(c$)
    c$ = "update kontakt set vid='" + id$ + "' where id='" + adrselect.kontselid + "'"
    Call form1.sqlqry(c$)
  Else
    'aus adresse
    Do
      kid$ = Left(trm(str$(Rnd)), 10)
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      c$ = "SELECT id FROM kontakt where id='" + kid$ + "'"
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    Loop Until stmp.EOF
    c$ = "insert into kontakt (id) values('" + kid$ + "');"
    Call form1.sqlqry(c$)
    cmd$ = "update kontakt set vid='" + id$ + "' where id ='" + kid$ + "'"
    Call form1.sqlqry(cmd$)
    c$ = "select * from adresse where id='" + neuwert + "'"
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      c$ = "update kontakt set name='" + trmvalidate(stmp!name) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set strasse='" + trmvalidate(stmp!strasse) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set ort='" + trmvalidate(stmp!ort) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set tel='" + trmvalidate(stmp!tel) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set fax='" + trmvalidate(stmp!fax) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set email='" + trmvalidate(stmp!email) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set handy='" + trmvalidate(stmp!handy) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set telfaxhandy='" + trm(stmp!telfaxhandy) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set url='" + trmvalidate(stmp!url) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set plz='" + trmvalidate(stmp!plz) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set lkz='" + trmvalidate(stmp!land) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set plzpostfach='" + trmvalidate(stmp!plzpostfach) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set postfach='" + trmvalidate(stmp!postfach) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      c$ = "update kontakt set postanrede='" + trmvalidate(stmp!postanrede) + "' where id='" + kid$ + "'": Call form1.sqlqry(c$)
      If trmvalidate(stmp!hinweise) <> "" Then
        knid$ = form1.newid("auftritthigru", "id", 40)
        c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values(" + _
           "'" + knid$ + "'," + _
           "'" + trm(id$ + kid$) + "'," + _
           "'Person'," + _
           "'Hinweise'," + _
           "'" + trm(stmp!hinweise) + "')"
        Call form1.sqlqry(c$)
        If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
          If kid$ <> "" Then
            c$ = "update auftritthigru set opt_kid='" + kid$ + "' where id='" + knid$ + "'"
            Call form1.sqlqry(c$)
          End If
        End If
Call form1.dbg2f(c$, "shwadrdetail", "Commad43_Click")
      End If
    End If
    If Not form1.isfieldmissing("opt_adresspool", "id") Then
        c$ = "select * from opt_adresspool where vid='" + trm(neuwert) + "' and kid='-1'"
        Set stmp = New ADODB.Recordset
        stmp.CursorLocation = adUseServer
        rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If rrr = 0 Then
          While Not stmp.EOF
            c$ = "insert into opt_adresspool (id,vid,kid,Beschreibung,Ort,Strasse,PLZ,Postfach,Land,Bundesland,PLZPostfach) values('" + vid$ + kid$ + trm(stmp!Beschreibung) + "','" + vid$ + "','" + kid$ + "',"
            c$ = c$ + "'" + trm(stmp!Beschreibung) + "',"
            c$ = c$ + "'" + trm(stmp!ort) + "',"
            c$ = c$ + "'" + trm(stmp!strasse) + "',"
            c$ = c$ + "'" + trm(stmp!plz) + "',"
            c$ = c$ + "'" + trm(stmp!postfach) + "',"
            c$ = c$ + "'" + trm(stmp!land) + "',"
            c$ = c$ + "'" + trm(stmp!Bundesland) + "',"
            c$ = c$ + "'" + trm(stmp!plzpostfach) + "')"
            Call form1.sqlqry(c$)
            stmp.MoveNext
          Wend
        End If
    End If
    
    c$ = "select * from anreden where kid='-1." + neuwert + "'"
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      c$ = "insert into anreden (id,kid,user,an,ab) values(" + _
           "'" + form1.newid("anreden", "id", 18) + "'," + _
           "'" + kid$ + "'," + _
           "'" + trm(stmp!user) + "'," + _
           "'" + trm(stmp!an) + "'," + _
           "'" + trm(stmp!Ab) + "')"
      Call form1.sqlqry(c$)
      stmp.MoveNext
    Wend
    c$ = "select * from auftritthigru where auftrittsid='" + neuwert + "'"
Call form1.dbg2f(c$, "shwadrdetail", "Commad43_Click (neuwert)")
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      knid$ = form1.newid("auftritthigru", "id", 40)
      c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values(" + _
           "'" + knid$ + "'," + _
           "'" + trm(id$ + kid$) + "'," + _
           "'" + trm(stmp!auftrittstyp) + "'," + _
           "'" + trm(stmp!feldname) + "'," + _
           "'" + trm(stmp!felddaten) + "')"
      Call form1.sqlqry(c$)
      If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
        If kid$ <> "" Then
          c$ = "update auftritthigru set opt_kid='" + kid$ + "' where id='" + knid$ + "'"
          Call form1.sqlqry(c$)
        End If
      End If
Call form1.dbg2f(c$, "shwadrdetail", "Commad43_Click")
      
      stmp.MoveNext
    Wend
    c$ = "select * from adresstyp where vid='" + neuwert + "' and kid='-1'"
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      c$ = "insert into adresstyp (id,vid,typ,wert,kid) values(" + _
           "'" + form1.newid("adresstyp", "id", 18) + "'," + _
           "'" + trm(id$) + "'," + _
           "'" + trm(stmp!typ) + "'," + _
           "'" + trm(stmp!wert) + "'," + _
           "'" + kid$ + "')"
      Call form1.sqlqry(c$)
      
      typ$ = trm(stmp!typ)
      If Left(typ$, 4) = "rel:" Then
        ityp$ = Mid(typ$, 5)
        invtyp = form1.getusersetting("inversrelation_" + ityp$, "")
        If invtyp <> "" Then
          invwert = form1.getkontaktnamebyid(kid$) + " {" + id$ + "}"
          sida$ = trm(stmp!wert)
          sidk$ = "-1"
          If InStr(trm(stmp!wert), "{") > 0 Then
            sid$ = trm(stmp!wert)
            sidp% = InStr(sid$, "{")
            sida$ = sid$
            sidk$ = trm(Left(sid$, sidp% - 1))
            sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
            sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
          End If
          c$ = "select * from adresstyp where vid='" + sida$ + "' and kid='" + sidk$ + "' and typ='rel:" + invtyp + "' and wert='" + invwert + "'"
          Set rtmp = New ADODB.Recordset
          rtmp.CursorLocation = adUseServer
          rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If rtmp.EOF Then
            form1.sqlqry _
             ( _
              "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + _
               sida$ + "','rel:" + invtyp + "','" + invwert + "','" + sidk$ + "')" _
             )
          End If
        End If
      End If
      
      stmp.MoveNext
    Wend
  End If
  Call refreshadrdetail(id$, "")
End If
MousePointer = 0

End Sub

Private Sub Command44_Click()
Dim i%, c$

i% = klist.ListIndex
If i% > klist.ListCount - 2 Then Exit Sub
Call savecheck
MousePointer = 11: DoEvents
form1.noalarms = True
c$ = "update kontakt set opt_kpos=" + trm(i% + 1) + " where id='" + idxlist.List(i%) + "'"
Call form1.sqlqry(c$)
c$ = "update kontakt set opt_kpos=" + trm(i%) + " where id='" + idxlist.List(i% + 1) + "'"
Call form1.sqlqry(c$)
Call refreshadrdetail(datf(0).text, klist.List(i%))
form1.noalarms = False
MousePointer = 0
End Sub

Private Sub Command45_Click()
Dim i%, c$

i% = klist.ListIndex
If i% < 1 Then Exit Sub
Call savecheck
MousePointer = 11: DoEvents
form1.noalarms = True
c$ = "update kontakt set opt_kpos=" + trm(i% - 1) + " where id='" + idxlist.List(i%) + "'"
Call form1.sqlqry(c$)
c$ = "update kontakt set opt_kpos=" + trm(i%) + " where id='" + idxlist.List(i% - 1) + "'"
Call form1.sqlqry(c$)
Call refreshadrdetail(datf(0).text, klist.List(i%))
form1.noalarms = False
MousePointer = 0
End Sub

Private Sub Command46_Click()
Dim i%, up$, cmd$, id$, kid$, stmp As ADODB.Recordset, neukwert, neukwertid, neuwert, neuawert, c$
Dim rrr, typ$, invtyp As String, ityp$
Dim sid$, sidp%, sida$, sidk$, rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command46_Click"
Call savecheck
id$ = datf(0).text
If id$ <> "" Then Call savecheck
Unload adrselect
DoEvents
Load adrselect
Call adrselect.sel_init("", "")
Call adrselect.SetFocus
Do
  DoEvents
Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
If adrselect.sel_brk() = 0 Then
  MousePointer = 11: DoEvents
  neukwertid = adrselect.get_kontselid()
  neukwert = adrselect.get_kontsel()
  neuwert = adrselect.sel_getselected()
  Unload adrselect
  If neukwertid <> "" Then
    Do
      id$ = word2bis(trm(neukwert)) + ", " + word1(trm(neukwert))
      i% = 0
      Do
        idn$ = id$
        If i% > 0 Then idn$ = id$ + "_" + trm(i%)
        i% = i% + 1
      Loop Until form1.getidbyid(idn$) = ""
      neuid$ = InputBox(transe("Neuer Sortiername:"), transe("Neue Adresse anlegen"), idn$)
    Loop Until form1.getidbyid(neuid$) = ""
    neuid = strrepl(trm(neuid), "/", "_")
    neuid = strrepl(neuid, "'", "´")
    If neuid$ <> "" Then
      If InStr(neuid$, "(") > 0 Then
        MsgBox transe("Unerlaubtes Zeichen in der ID.")
        Exit Sub
      End If
      s$ = "insert into adresse (id,name) values('" & neuid$ & "','" & neukwert & "')"
      Call form1.sqlqry(s$)
      If Not form1.isfieldmissing("opt_adresspool", "id") Then
        c$ = "select * from opt_adresspool where kid='" + neukwertid + "'"
        Set stmp = New ADODB.Recordset
        stmp.CursorLocation = adUseServer
        rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If rrr = 0 Then
          While Not stmp.EOF
            c$ = "insert into opt_adresspool (id,vid,kid,Beschreibung,Ort,Strasse,PLZ,Postfach,Land,Bundesland,PLZPostfach) values('" + neuid + "-1" + trm(stmp!Beschreibung) + "','" + neuid + "','-1',"
            c$ = c$ + "'" + trm(stmp!Beschreibung) + "',"
            c$ = c$ + "'" + trm(stmp!ort) + "',"
            c$ = c$ + "'" + trm(stmp!strasse) + "',"
            c$ = c$ + "'" + trm(stmp!plz) + "',"
            c$ = c$ + "'" + trm(stmp!postfach) + "',"
            c$ = c$ + "'" + trm(stmp!land) + "',"
            c$ = c$ + "'" + trm(stmp!Bundesland) + "',"
            c$ = c$ + "'" + trm(stmp!plzpostfach) + "')"
            Call form1.sqlqry(c$)
            stmp.MoveNext
          Wend
        End If
      End If
      c$ = "select * from kontakt where id='" + neukwertid + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
        c$ = "update adresse set tel='" + trm(stmp!tel) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set fax='" + trm(stmp!fax) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set handy='" + trm(stmp!handy) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set email='" + trm(stmp!email) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set telfaxhandy='" + trm(stmp!telfaxhandy) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set hinweise='" + trm(stmp!Position) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set Strasse='" + trm(stmp!strasse) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set Ort='" + trm(stmp!ort) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set url='" + trm(stmp!url) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set PLZPostfach='" + trm(stmp!plzpostfach) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set PLZ='" + trm(stmp!plz) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set Postfach='" + trm(stmp!postfach) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set Postanrede='" + trm(stmp!postanrede) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set PLZPostfach='" + trm(stmp!plzpostfach) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
        c$ = "update adresse set Land='" + trm(stmp!lkz) + "' where id='" + neuid$ + "'": Call form1.sqlqry(c$)
      End If
      c$ = "select * from adresstyp where kid='" + neukwertid + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
      While Not stmp.EOF
        id$ = form1.newid("adresstyp", "id", 15)
        c$ = "insert into adresstyp (id,vid,typ,wert,kid) values('" + id$ + "','" + neuid$ + "','" + _
             trm(stmp!typ) + "','" + trm(stmp!wert) + "','-1')"
        Call form1.sqlqry(c$)
'Debug.Print stmp!typ
        typ$ = trm(stmp!typ)
        If Left(typ$, 4) = "rel:" Then
          ityp$ = Mid(typ$, 5)
          invtyp = form1.getusersetting("inversrelation_" + ityp$, "")
          If invtyp <> "" Then
            invwert = neuid$
            sida$ = trm(stmp!wert)
            sidk$ = "-1"
            If InStr(trm(stmp!wert), "{") > 0 Then
              sid$ = trm(stmp!wert)
              sidp% = InStr(sid$, "{")
              sida$ = sid$
              sidk$ = trm(Left(sid$, sidp% - 1))
              sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
              sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
            End If
            c$ = "select * from adresstyp where vid='" + sida$ + "' and kid='" + sidk$ + "' and typ='rel:" + invtyp + "' and wert='" + neuid$ + "'"
            Set rtmp = New ADODB.Recordset
            rtmp.CursorLocation = adUseServer
            rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If rtmp.EOF Then
              form1.sqlqry _
               ( _
                "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + _
                 sida$ + "','rel:" + invtyp + "','" + invwert + "','" + sidk$ + "')" _
               )
            End If
          End If
        End If

        stmp.MoveNext
      Wend
      End If
      c$ = "select * from auftritthigru where auftrittsid='" + neuwert + neukwertid + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
      While Not stmp.EOF
        id$ = form1.newid("auftritthigru", "id", 15)
        c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,Feldname,Felddaten) values('" + id$ + "','" + neuid$ + "','" + _
             trm(stmp!auftrittstyp) + "','" + trm(stmp!feldname) + "','" + trm(stmp!felddaten) + "')"
        Call form1.sqlqry(c$)
        stmp.MoveNext
      Wend
      End If
      c$ = "select * from anreden where kid='" + neukwertid + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
      While Not stmp.EOF
        id$ = form1.newid("anreden", "id", 15)
        c$ = "insert into anreden (id,kid,user,An,Ab) values('" + id$ + "','-1." + neuid$ + "','" + _
             trm(stmp!user) + "','" + trm(stmp!an) + "','" + trm(stmp!Ab) + "')"
        Call form1.sqlqry(c$)
        stmp.MoveNext
      Wend
      End If
      c$ = "select * from dochist where kontakt='" + neukwertid + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
      While Not stmp.EOF
        id$ = form1.newid("dochist", "id", 15)
        c$ = "insert into dochist (id,adresse,kontakt,docname,owner,Betreff,Memoinhalt,doctyp) values('" + id$ + "','" + neuid$ + "','-1','" + _
             trm(stmp!docname) + "','" + trm(stmp!Owner) + "','" + trm(stmp!betreff) + "','" + trm(stmp!Memoinhalt) + "','" + trm(stmp!doctyp) + "')"
        Call form1.sqlqry(c$)
        stmp.MoveNext
      Wend
      End If

      Call refreshadrdetail(neuid$, "")
    End If
  End If
End If
MousePointer = 0: DoEvents
End Sub

Private Sub Command47_Click()
Dim i%

Combo1.text = transe("keine Liste")
DoEvents
For i% = 0 To List3.ListCount - 1
  If List3.Selected(i%) Then List3.Selected(i%) = False
Next i%
Call Command15_Click
End Sub

Private Sub Command48_Click()
Dim c$, r As ADODB.Recordset, cid$

If knt_sav.Enabled Then Call savecheck
If form1.isfieldmissing("opt_adresspool", "id") Then
  MsgBox ("Die erforderliche Tabelle opt_adresspool fehlt." + vbCrLf + "Bitte kontaktieren Sie den Support.")
  Exit Sub
End If
gd2.ListItems.Clear
cid$ = datf(0).text
If cid$ = "" Then Exit Sub
Call savecheck
gd2.Visible = True

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c$ = "select id,Beschreibung,Ort,Strasse from opt_adresspool where vid='" + cid$ + "' and kid='-1' order by Beschreibung"
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If r.EOF Then
  Call svadr("Standard")
  Set lvitem = gd2.ListItems.add(, , trm("Standard"))
  lvitem.SubItems(1) = trm(datf(3).text)
  lvitem.SubItems(2) = trm(datf(2).text)
Else
  While Not r.EOF
    Set lvitem = gd2.ListItems.add(, , trm(r!Beschreibung))
    lvitem.SubItems(1) = trm(r!ort)
    lvitem.SubItems(2) = trm(r!strasse)
    lvitem.SubItems(3) = trm(r!id)
    r.MoveNext
  Wend
End If
End Sub

Private Sub Command49_Click()
Dim c$, r As ADODB.Recordset, cid$

If form1.isfieldmissing("opt_adresspool", "id") Then
  MsgBox ("Die erforderliche Tabelle opt_adresspool fehlt." + vbCrLf + "Bitte kontaktieren Sie den Support.")
  Exit Sub
End If
gd3.ListItems.Clear
cid$ = datf(0).text
If cid$ = "" Then Exit Sub
kid$ = kdat(0).text
If kid$ = "" Then Exit Sub
Call savecheck
gd3.Visible = True

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c$ = "select id,Beschreibung,Ort,Strasse from opt_adresspool where vid='" + cid$ + "' and kid='" + kid$ + "' order by Beschreibung"
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If r.EOF Then
  Call svkadr("Standard")
  Set lvitem = gd3.ListItems.add(, , trm("Standard"))
  lvitem.SubItems(1) = trm(kadat(3).text)
  lvitem.SubItems(2) = trm(datf(0).text)
Else
  While Not r.EOF
    Set lvitem = gd3.ListItems.add(, , trm(r!Beschreibung))
    lvitem.SubItems(1) = trm(r!ort)
    lvitem.SubItems(2) = trm(r!strasse)
    lvitem.SubItems(3) = trm(r!id)
    r.MoveNext
  Wend
End If

End Sub

Public Sub Command5_Click()
Dim i%, up$, cmd$, antxt$, abtxt$, resoredata$, fn$, id$
Dim restoredata$, c$, fld$, anid$, fldns$

'd2infile = "shwAdrDetail": d2insub = "Command5_Click"
id$ = kdat(0).text
If id$ <> "" Then
If cadrpbez.text <> "Standard" Then
  up$ = trm(InputBox(transe("Sie speichern nicht die Standardadresse." + vbCrLf + "Die jetzt eingetragene Adresse wird die Standardadresse." + vbCrLf + "Bestätigen Sie bitte mit JA"), transe("Adresse unklar.")))
  If LCase(trm(up$)) <> "ja" Then Exit Sub
  Call svkadr("Standard")
End If

If form1.getusersetting("adressenstandautomatisch", "ja") = "ja" Then
  datf(8) = trm(Date) + " " + trm(Time)
End If
up$ = form1.getusersetting("adresshistorie", "nein")
If up$ = "ja" Or up$ = "erweitert" Then
  'If up$ = "erweitert" Then
    restoredata$ = "N/A"
    fn$ = form1.s0dir() + "\tmp\" + form1.medienname(datf(0).text) + "." + id$
    i% = FreeFile
    'Open fn$ For Output As #i%: Close #i%
    c$ = "delete from dochist where adresse='" + datf(0).text + "' and kontakt='" + id$ + "' and doctyp='" + transe("Datenänderung") + "';"
    Call form1.sqlqry(c$)
    c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
            form1.newid("dochist", "id", 19) & "','" & datf(0).text & "','" + id$ + "','" + fn$ & "','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','" + transe("geändert") & "','" + transe("Datenänderung") + "')"
    Call form1.sqlqry(c$)
  'End If
End If

fldns$ = "id"
For i% = 1 To nfldsk
  fldns$ = fldns$ + "," + form1.sqla.TableDefs("kontakt").Fields(i%).name
Next i%

For i% = 0 To nfldsk
  up$ = cut_d1(fldns$, ","): fldns$ = cut_d2bis(fldns$, ",")
  If i% > 0 Then
    cmd$ = "update kontakt set " + up$ + " = '" + strrepl(kdat(i%).text, "'", "´") + "' where id='" & id$ & "'"
    Call form1.sqlqry(cmd$)
  End If
Next i%
form1.noalarms = True
For i% = 0 To 5
  Select Case i%
    Case 0: fld$ = "strasse"
    Case 1: fld$ = "lkz"
    Case 2: fld$ = "plz"
    Case 3: fld$ = "ort"
    Case 4: fld$ = "PLZPostfach"
    Case 5: fld$ = "Postfach"
  End Select
  cmd$ = "update kontakt set " & fld$ & "='" & trmvalidate(kadat(i%).text) & " ' where id='" & id$ & "'"
  Call form1.sqlqry(cmd$)
Next i%
cmd$ = "update kontakt set postanrede='" & trmvalidate(postanredek.text) & " ' where id='" & id$ & "'"
Call form1.sqlqry(cmd$)
If trm(datf(0).text) <> "" Then Call form1.chkallnums(datf(0).text, id$, "email", kdat(5).text)
If Not form1.isfieldmissing("kontakt", "opttel") Then
  If optktel.text <> "" Then
    cmd$ = "update kontakt set opttel='" & optktel.text & " ' where id='" & id$ & "'"
    Call form1.sqlqry(cmd$)
  End If
End If

form1.sqlqry ("delete from anreden where kid='" + id$ + "' and user='" + form1.anredeuser$ + "'")

antxt$ = trm(Anrede.text)
abtxt$ = trm(Abrede.text)
anid$ = form1.newid("anreden", "id", 18)
c$ = "insert into anreden (id,kid,user) values('" + anid$ + "','" + id$ + "','" + form1.anredeuser$ + "')"
Call form1.sqlqry(c$)
If Len(antxt$) > 0 Then form1.sqlqry ("update anreden set an='" & antxt$ & "' where id='" + anid$ + "'")
If Len(abtxt$) > 0 Then form1.sqlqry ("update anreden set ab='" & abtxt$ & "' where id='" + anid$ + "'")
If form1.getusersetting("systemanredensetzen", "nein") = "ja" Then
  form1.sqlqry ("delete from anreden where kid='" + id$ + "' and user='system'")
  anid$ = form1.newid("anreden", "id", 18)
  c$ = "insert into anreden (id,kid,user) values('" + anid$ + "','" + id$ + "','system')"
  Call form1.sqlqry(c$)
  If Len(antxt$) > 0 Then
    form1.sqlqry ("update anreden set an='" & antxt$ & "' where id='" + anid$ + "'")
  End If
  If Len(abtxt$) > 0 Then form1.sqlqry ("update anreden set ab='" & abtxt$ & "' where id='" + anid$ + "'")
End If
form1.noalarms = False
End If 'kontakt gewählt?
Call form1.kontakt2cloud(id$)
Call Command4_Click
Call form1.combo1_Change
If klist.ListCount > 0 Then klist.ListIndex = 0
End Sub

Private Sub Command50_Click()
    
    Load kc
    kc.selct(2).Clear
    If idshow.Caption <> "" Then
      kc.selct(2).AddItem idshow.Caption
      For i% = 0 To klist.ListCount - 1
        kc.selct(2).AddItem klist.List(i%) + " {" + idshow.Caption + "}"
      Next i%
    End If
    Call kc.settag0(Date)
    Call kc.Command1_Click
    On Error Resume Next
    Call kc.SetFocus
    Call k3.SetFocus
    On Error GoTo 0

End Sub

Private Sub Command51_Click()

If datf(7).Height < 1000 Then
  datf(7).Top = 840
  datf(7).Width = 8655
  datf(7).Height = 4965
  Command51.Caption = "-"
Else
  datf(7).Top = 5400
  datf(7).Width = 3375
  datf(7).Height = 765
  Command51.Caption = "+"
End If

End Sub


Private Sub Command52_Click()
Dim id$, c$, ask%, ml$

ml$ = trm(datf(6).text)
If ml$ = "" Then Exit Sub
id$ = datf(0).text
If id$ = "" Then Exit Sub
ask% = MsgBox(transe("E-Mailadresse löschen") + ": " & ml$ + vbCrLf + transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("E-Mail") & " " & ml$ & " " & transe("löschen") & "?")
If ask% <> vbYes Then Exit Sub
c$ = "delete FROM opt_allenummern where vid='" + id$ + "' and kid='-1' and numtyp='email' and num='" + ml$ + "'"
Call form1.sqlqry(c$)
datf(6).text = ""
Call Command4_Click
End Sub

Private Sub Command53_Click()
Dim id$, c$, ask%, ml$

ml$ = trm(kdat(5).text)
If ml$ = "" Then Exit Sub
id$ = datf(0).text
If id$ = "" Then Exit Sub
kid$ = kdat(0).text
If kid$ = "" Or kid$ = "-1" Then Exit Sub
ask% = MsgBox(transe("E-Mailadresse löschen") + ": " & ml$ + vbCrLf + transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("E-Mail") & " " & ml$ & " " & transe("löschen") & "?")
If ask% <> vbYes Then Exit Sub
c$ = "delete FROM opt_allenummern where vid='" + id$ + "' and kid='" + kid$ + "' and numtyp='email' and num='" + ml$ + "'"
Call form1.sqlqry(c$)
kdat(5).text = ""
Call Command5_Click

End Sub

Private Sub Command6_Click()
Dim i%, up$, cmd$, rtmp As ADODB.Recordset, c$, adlist$, ri$, id$, rrr
Dim knam$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command6_Click"
Call savecheck
'Command6.Visible = False
Check1.value = 1
id$ = kdat(0).text
If id$ = "" Then Exit Sub
knam$ = kdat(2).text
c$ = "delete from anreden where kid='" + id$ + "." + trm(datf(0).text) + "'"
Call form1.sqlqry(c$)
c$ = "update dochist set kontakt='-1' where kontakt='" + id$ + "' and adresse='" + trm(datf(0).text) + "'"
Call form1.sqlqry(c$)
adlist$ = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT typ,wert FROM adresstyp where vid='" & datf(0).text & "' and kid='" & id$ + "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  While Not rtmp.EOF
    If form1.isoftype(datf(0).text, rtmp!typ) = "-1" Then
      If adlist$ <> "" Then adlist$ = adlist$ & ", "
      adlist$ = adlist$ & rtmp!typ
      ri$ = form1.newid("adresstyp", "id", 20)
      Call form1.sqlqry("delete from auftritthigru where auftrittsid='" + trm(datf(0).text) + "' and auftrittstyp='" + trm(rtmp!typ) + "'")
      Call form1.sqlqry("insert into adresstyp (id,vid,typ,kid) values('" & ri$ & "','" & datf(0).text & "','" + rtmp!typ + "','-1')")
      If trm(rtmp!wert) <> "" Then
        c$ = "update adresstyp set wert='" & trm(rtmp!wert) & "' where id='" & ri$ & "'"
        Call form1.sqlqry(c$)
      End If
      c$ = "update auftritthigru set auftrittsid='" + trm(datf(0).text) + "' where auftrittsid='" + trm(datf(0).text) + id$ + "' and auftrittstyp='" + trm(rtmp!typ) + "'"
      Call form1.sqlqry(c$)
      If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
        c$ = "update auftritthigru set opt_kid='' where auftrittsid='" + trm(datf(0).text) + "' and auftrittstyp='" + trm(rtmp!typ) + "'"
        Call form1.sqlqry(c$)
      End If
    Else
      c$ = "delete from auftritthigru where auftrittsid='" + trm(datf(0).text) + id$ + "' and auftrittstyp='" + trm(rtmp!typ) + "'"
      Call form1.sqlqry(c$)
    End If
    rtmp.MoveNext
  Wend
End If


cmd$ = "delete from kontakt where id='" + id$ + "'"
form1.sqlqry (cmd$)
cmd$ = "delete from adresstyp where vid='" + datf(0).text + "' and kid='" + id$ + "';"
form1.sqlqry (cmd$)
cmd$ = "delete from adresstyp where wert='" + knam$ + " {" + trm(datf(0).text) + "}';"
form1.sqlqry (cmd$)
cmd$ = "delete from auftritthigru where auftrittsid='" + datf(0).text + id$ + "'"
form1.sqlqry (cmd$)
cmd$ = "delete from anreden where kid='" + id$ + "'"
form1.sqlqry (cmd$)
If Not form1.isfieldmissing("opt_allenummern", "vid") Then
  c$ = "delete FROM opt_allenummern where vid='" + datf(0).text + "' and kid='" + id$ + "'"
  Call form1.sqlqry(c$)
End If
If adlist$ <> "" Then
  If form1.getuserid() <> "www" Then
    MsgBox (transe("Die Adresse wurde diesen Kategorien zugeordnet:") + vbCrLf & adlist$)
  End If
End If
Call Command4_Click
Call form1.combo1_Change

End Sub

Private Sub Command7_Click()
Dim i%, up$, cmd$, stmp As ADODB.Recordset, id$, c$, rrr

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "Command7_Click"
id$ = datf(0).text
If id$ = "" Then Exit Sub
Call savecheck
Do
  id$ = Left(trm(str$(Rnd)), 10)
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
  c$ = "SELECT id FROM kontakt where id='" + id$ + "'"
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Loop Until stmp.EOF

atp$ = form1.grantadrtyp()
up$ = "insert into kontakt (id,"
For i% = 1 To nfldsk
  up$ = up$ + form1.sqla.TableDefs("kontakt").Fields(i%).name + ","
Next i%
up$ = Left$(up$, Len(up$) - 1) + ") values('" + id$ + "',"
'up$ = "insert into kontakt values('" + id$ + "',"
For i% = 1 To nfldsk
  Select Case i%
    Case 2: kdat(i%).text = "Neuer Kontakt"
    Case 1: kdat(i%).text = datf(0).text
    Case Else: kdat(i%).text = ""
  End Select

  If Len(kdat(i%).text) = 0 Then
    up$ = up$ + "NULL,"
  Else
    up$ = up$ + "'" + kdat(i%).text + "',"
  End If
Next i%
up$ = Left$(up$, Len(up$) - 1)
cmd$ = up$ + ")"
Call form1.sqlqry(cmd$)

Call Command4_Click
Call form1.combo1_Change
For i% = 0 To klist.ListCount - 1
  If klist.List(i%) = "Neuer Kontakt" Then
    klist.ListIndex = i%
    DoEvents
    i% = klist.ListCount
    kdat(2).SetFocus
    Call addtyp("Person")
    If atp$ <> "" Then
      List1.AddItem transe(atp$)
      Call addtyp(atp$)
    End If
  End If
Next i%

End Sub

Public Sub Command8_Click()

'd2infile = "shwAdrDetail": d2insub = "Command8_Click"
Load adrtypselector
adrtypselector.Visible = True
Call adrtypselector.SetFocus
End Sub

Public Sub Command9_Click()
Dim typ$, vid$, wert$, kontakt$, hid$, c$
Dim cmd$, sidp%, sida$, sidk$

'd2infile = "shwAdrDetail": d2insub = "Command9_Click"
Unload bezlist
If Not List1b.Visible Then

If List1.ListIndex < 0 Then Exit Sub

typ$ = transo(List1.List(List1.ListIndex))
wert$ = ""
If InStr(typ$, ":") Then
  wert$ = Mid$(typ$, InStr(typ$, ":") + 1)
  typ$ = Left$(typ$, InStr(typ$, ":") - 1)
End If
vid$ = datf(0).text
If vid$ = "" Then Exit Sub
If klist.ListIndex >= 0 Then
  kontakt$ = idxlist.List(klist.ListIndex)
Else
  kontakt$ = "-1"
End If
ask% = MsgBox("Kategorie " & typ$ + vbCrLf + transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Kategorie") & " " & transe(typ$) & " " & transe("löschen") & "?")
If ask% <> vbYes Then Exit Sub
c$ = "delete from adresstyp where vid='" + vid$ + "' and typ='" + typ$ + "' and kid='" + kontakt$ + "'"
If Not form1.sqlqry(c$) Then
  Exit Sub
End If
List1.RemoveItem List1.ListIndex
If form1.getusersetting("adresstypwechsellog", "nein") = "ja" Then
     hid$ = form1.newid("dochist", "id", 19)
     c$ = "insert into dochist (id,adresse,kontakt,docname,doctyp,erstellt,owner,betreff,doctyp) values('" & _
            hid$ & "','" & vid$ & "','" + kontakt$ + "','Gruppenwechsel','Gruppenwechsel','" & _
           datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','entfernt aus AdrGrp " + typ$ + "','Emaileingang')"
     Call form1.sqlqry(c$)
End If
Call rlist4

Else


typ$ = List1b.List(List1b.ListIndex)
wert$ = ""
If InStr(typ$, ":") Then
  wert$ = Mid$(typ$, InStr(typ$, ":") + 1)
  typ$ = Left$(typ$, InStr(typ$, ":") - 1)
End If
vid$ = datf(0).text
If vid$ = "" Then Exit Sub
If klist.ListIndex >= 0 Then
  kontakt$ = idxlist.List(klist.ListIndex)
Else
  kontakt$ = "-1"
End If
c$ = "delete from adresstyp where vid='" + vid$ + "' and typ='rel:" + typ$ + "' and kid='" + kontakt$ + "'"
If wert$ <> "" Then c$ = c$ + " and wert='" + trm(wert$) + "'"
form1.sqlqry (c$)
invtyp = form1.getusersetting("inversrelation_" + typ$, "")
invwert = vid$
If kontakt$ <> "-1" And kontakt$ <> "'-1'" Then
  invwert = form1.get_kontaktname_by_id(kontakt$) + " {" + invwert + "}"
End If
If invtyp <> "" Then
  sidk$ = "-1"
  sida$ = wert$
  If InStr(wert$, "{") > 0 Then
    sid$ = wert$
    sidp% = InStr(sid$, "{")
    sidk$ = trm(Left(sid$, sidp% - 1))
    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  End If
  cmd$ = "delete from adresstyp where vid='" + trm(sida$) + "' and kid='" + sidk$ + "' and typ='rel:" + invtyp + "' and wert='" + invwert + "'"
'Debug.Print cmd$
   Call form1.sqlqry(cmd$)
End If
List1b.RemoveItem List1b.ListIndex
If List1b.ListCount > 0 Then List1b.ListIndex = 0


End If

End Sub

Private Sub datf_Change(Index As Integer)

'd2infile = "shwAdrDetail": d2insub = "datf_Change"
If Index <> 0 Then
  Command4.Enabled = True
  adr_sav.Enabled = True
  BackColor = form1.dirtycolor()
  'telfaxhandy
  If Index = 4 Or Index = 5 Or Index = 9 Then
    datf(12).text = onlynums(datf(4).text) + " " + onlynums(datf(5).text) + " " + onlynums(datf(9).text) + " "
    If LCase(form1.getusersetting("opttel", "Tel")) = "tel" Then datf(12).text = datf(12).text + " " + onlynums(opttel.text)
  End If
Else
  If srchit% <> 0 Then Combo3.text = datf(0).text
  idshow.Caption = datf(0)
End If

End Sub

Private Sub datf_DblClick(Index As Integer)
Dim w$, z$, i%, f$

'd2infile = "shwAdrDetail": d2insub = "datf_DblClick"
If Index = 13 Or Index = 3 Then Exit Sub

If Index = 7 Then
  Load memoview
  Call memoview.settext(datf(Index).text)
  Exit Sub
End If

If Index = 10 Then
  Unload frmBrowser
  DoEvents
  frmBrowser.StartingAddress = datf(Index).text
  Load frmBrowser
  Exit Sub
End If

If form1.darf_ich_sprechen() = True Then Call spknum(datf(Index).text())

End Sub

Public Sub spknum(w$)
Dim i%, z$

  For i% = 1 To Len(w$)
    z$ = Mid$(w$, i%, 1)
    If isdigit(z$) > 0 Then
      f$ = form1.wavdir() & "\" & z$ & ".wav"
      If exist(f$) <> 0 Then
        Call sndPlaySound(f$, SND_SYNC)
      End If
    End If
    DoEvents
  Next i%

End Sub

Private Sub datf_GotFocus(Index As Integer)

'd2infile = "shwAdrDetail": d2insub = "datf_GotFocus"
prv$ = datf(Index).text
If knt_sav.Enabled Then Call savecheck

End Sub

Private Sub datf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

'd2infile = "shwAdrDetail": d2insub = "datf_KeyDown"
If Index = 3 Or Index = 13 Or Index = 14 Then Call esccntinc(KeyCode)
If Index = 2 Then Call esccntinc2(KeyCode)

End Sub

Sub esccntinc(kc As Integer)
'd2infile = "shwAdrDetail": d2insub = "esccntinc"
If kc = 27 Then
  esccnt = esccnt + 1
  If esccnt > 1 Then
    esccnt = 0
    Call Label30_DblClick
  End If
Else
  esccnt = 0
End If

End Sub

Sub esccntinc2(kc As Integer)
'd2infile = "shwAdrDetail": d2insub = "esccntinc2"
If kc = 27 Then
  esccnt2 = esccnt2 + 1
  If esccnt2 > 1 Then
    esccnt2 = 0
    Call Label3_DblClick
  End If
Else
  esccnt2 = 0
End If

End Sub

Private Sub datf_LostFocus(Index As Integer)

datf(Index).text = strrepl(datf(Index).text, "'", "´")

End Sub

Private Sub datf_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fn$, nf%, rrr

'd2infile = "shwAdrDetail": d2insub = "datf_OLEDragDrop"
If idshow.Caption <> "" Or datf(1).text <> "" Then Exit Sub
If Data.GetFormat(vbCFFiles) Then
  nf% = 1
  Do
    On Error Resume Next
    fn$ = Data.Files(nf%)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      If LCase(Right(fn$, 4)) = ".csv" Then
        Call importcsv(fn$)
      End If
      DoEvents
    End If
    nf% = nf% + 1
  Loop Until rrr <> 0
End If
End Sub

Private Sub dok_brief_Click()
Call Command3_Click
End Sub

Private Sub dok_eml_Click()
If klist.ListIndex < 0 Then
  Call Command17_Click(0)
Else
  Call Command17_Click(1)
End If

End Sub

Private Sub dok_fax_Click()
Call Command2_Click
End Sub

Private Sub dok_kont_Click()
Call Command10_Click
End Sub

Private Sub dok_memo_Click()
Call Command16_Click
End Sub

Private Sub Form_Load()
Dim klrv%, k1lrv%, colHeader, s%, i%, tr As String, rrr, f$
Dim dbpara$, adopara$, cf$

'd2infile = "shwAdrDetail": d2insub = "Form_Load"
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)

rlist2icalmode = False
nodbupd = False
gd1.Visible = False
gd1.View = lvwReport
gd2.Visible = False
gd2.View = lvwReport
gd3.Visible = False
gd3.View = lvwReport
esccnt = 0
stcky_igno = False
esckcnt = 0
gd1upd% = 1
List1b.Top = List1.Top
List1b.Left = List1.Left
List1b.Visible = False
Label4.ForeColor = form1.lnkcolor
Label52.Caption = transe("inkl. Kontakte")
Command47.Caption = transe("&alle zeigen")
Label5.ForeColor = form1.lnkcolor
Label6.ForeColor = form1.lnkcolor
Label11.ForeColor = form1.lnkcolor
Label10.ForeColor = form1.lnkcolor
Label24.ForeColor = form1.lnkcolor
Label27.ForeColor = form1.lnkcolor
Label28.ForeColor = form1.lnkcolor
prio.Enabled = False
prio.Visible = False
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:load")
Call form1.dbg2f("shwAdrDetail:load")
shwAdrDetail.Caption = form1.inmylanguage("Adressen - Agencyprof")
kadat(0).ToolTipText = form1.inmylanguage("Postfach")
kadat(1).ToolTipText = form1.inmylanguage("Postfach")
postf.ToolTipText = form1.inmylanguage("Postfach")
icalconf.Caption = form1.inmylanguage("Konfig")
icalconf.ToolTipText = form1.inmylanguage("Ausgewählte Termine als iCalendar exportieren")
Command34.Caption = "vCard"
Command34.ToolTipText = form1.inmylanguage("Adresse (oder ausgewählten Kontakt) als vCard versenden, Weitergabe der Handynummer einstellbar")
icalex.Caption = "-->iCAL"
icalex.ToolTipText = form1.inmylanguage("Ausgewählte Termine als iCalendar exportieren")
Command33.ToolTipText = form1.inmylanguage("Adresse in die Zwischenablage kopieren")
Command32.ToolTipText = form1.inmylanguage("Adresse als Dokument")
suchen.ToolTipText = "Zeige in Karte"
Command31.ToolTipText = form1.inmylanguage("per Email an Agencyprof")
Command30.Caption = form1.inmylanguage("Saalplan")
Command30.ToolTipText = form1.inmylanguage("Neue Biographie anlegen")
Command37.ToolTipText = form1.inmylanguage("Adresse kopieren")
Command46.ToolTipText = form1.inmylanguage("Neue Adresse aus einem Kontakt erstellen")
Command43.ToolTipText = form1.inmylanguage("Neuen Kontakt aus einer Adresse erstellen")
Command29.ToolTipText = form1.inmylanguage("... " + transe("im Explorer öffnen"))
opttel.ToolTipText = ""
optktel.ToolTipText = ""
Command28.Caption = "?"
Command27.Caption = form1.inmylanguage("Bühnenplan")
Command27.ToolTipText = form1.inmylanguage("Neue Biographie anlegen")
kdat(9).ToolTipText = form1.inmylanguage("Position")
usempth.Caption = form1.inmylanguage("Medienpfad")
kdat(8).ToolTipText = form1.inmylanguage("URL Kontaktperson")
Command26.ToolTipText = form1.inmylanguage("Neuen Termin anlegen")
Command23.Caption = form1.inmylanguage("neue Bio")
Command23.ToolTipText = form1.inmylanguage("Neue Biographie anlegen")
Command1.ToolTipText = form1.inmylanguage("Formular schliessen")
Command24.Caption = form1.inmylanguage("Verzeichnis öffnen")
Command24.ToolTipText = form1.inmylanguage("... " + transe("im Explorer öffnen"))
List9.ToolTipText = form1.inmylanguage("weitere Druckvorlagen")
Command25.ToolTipText = form1.inmylanguage("Wiedervorlage")
List7.ToolTipText = form1.inmylanguage("Diese Mediadateien sind hinterlegt")
Command22.Caption = ">>"
Command21.Caption = "<<"
Command20.Caption = "<"
Command19.Caption = ">"
Command18.Caption = form1.inmylanguage("Media Dateien")
Command18.ToolTipText = form1.inmylanguage("Medien anzeigen/ verbergen")
Command17(0).Caption = "@"
Command17(0).ToolTipText = form1.inmylanguage("Email senden an Kontaktperson")
Command17(1).Caption = "@"
Command17(1).ToolTipText = form1.inmylanguage("Email senden an Adresse")
Command16.ToolTipText = form1.inmylanguage("Memo schreiben")
Command52.ToolTipText = form1.inmylanguage(transe("Diese E-Mailadresse löschen"))
Command53.ToolTipText = form1.inmylanguage(transe("Diese E-Mailadresse löschen"))
Text1.text = "730"
Text1.ToolTipText = form1.inmylanguage("Wieviel Tage zurück")
Text2.text = "365"
Text2.ToolTipText = form1.inmylanguage("Wieviel Tage in die Zukunft")
List4.ToolTipText = form1.inmylanguage("Zusatzfelder je nach gewählten Kategorien")
Command15.Caption = "Go !"
Command50.ToolTipText = form1.inmylanguage("Kalender öffnen")
Command15.ToolTipText = form1.inmylanguage("Zeige Verknüpfte Termine")
List3.ToolTipText = form1.inmylanguage("Kategorien der Termine")
List2.ToolTipText = form1.inmylanguage("Liste verknüpfter Termine")
Command14.Caption = form1.inmylanguage("Zusatz-Infos")
Command14.ToolTipText = form1.inmylanguage("Zusatz-Informationen anzeigen/ verbergen")
Check2.ToolTipText = form1.inmylanguage("Zum Löschen deaktivieren")
Command13.ToolTipText = form1.inmylanguage("Lösche gesamten Datensatz")
Command12.Caption = form1.inmylanguage("Abwahl")
Command12.ToolTipText = form1.inmylanguage("Keine Kontaktperson auswählen")
kdat(7).ToolTipText = form1.inmylanguage("Handy-Nummer Kontaktperson")
Command11.ToolTipText = form1.inmylanguage("Neue Adresse anlegen")
Command10.Caption = form1.inmylanguage("Kontakt - Historie")
Command10.ToolTipText = form1.inmylanguage("Bisherige Kontakte zeigen")
Command9.Caption = "-"
Command9.ToolTipText = form1.inmylanguage("Kategorie hier entfernen")
Command8.Caption = "+"
Command8.ToolTipText = form1.inmylanguage("Kategorie hinzufügen")
List1.ToolTipText = form1.inmylanguage("Hier angewandte Kategorien")
Check1.ToolTipText = form1.inmylanguage("Zum Löschen deaktivieren")
Command7.ToolTipText = form1.inmylanguage("Neuen Kontakt anlegen")
Command6.ToolTipText = form1.inmylanguage("Löschen Kontakt")
Command5.ToolTipText = form1.inmylanguage("Kontakt speichern")
kdat(6).ToolTipText = form1.inmylanguage("Email Kontaktperson")
kdat(5).ToolTipText = form1.inmylanguage("Fax Kontaktperson")
datf(11).ToolTipText = form1.inmylanguage("Adress-Kürzel")
Command4.ToolTipText = form1.inmylanguage("Speichern")
kdat(4).ToolTipText = form1.inmylanguage("Telefon Kontaktperson")
kdat(3).ToolTipText = form1.inmylanguage("Name Kontaktperson")
Command3.ToolTipText = form1.inmylanguage("Brief schreiben")
Command2.Caption = form1.inmylanguage("&Fax")
Command2.ToolTipText = form1.inmylanguage("Fax schreiben")
Command3.ToolTipText = form1.inmylanguage("Brief schreiben")
datf(10).ToolTipText = form1.inmylanguage("Internet-Adresse")
datf(9).ToolTipText = form1.inmylanguage("Handy-Nummer")
datf(8).ToolTipText = form1.inmylanguage("Datum der letzte Änderung")
datf(7).ToolTipText = form1.inmylanguage("Bemerkungen - Doppelklick zur vergrößerten Ansicht")
datf(6).ToolTipText = form1.inmylanguage("Email-Adresse")
datf(5).ToolTipText = form1.inmylanguage("Fax-Nummer")
datf(4).ToolTipText = form1.inmylanguage("Telefon")
datf(3).ToolTipText = form1.inmylanguage("Land Postleitzahl Ort")
datf(2).ToolTipText = form1.inmylanguage("Straße/ Postfach")
datf(1).ToolTipText = form1.inmylanguage("Namen der Firma")
altbvorl.ToolTipText = form1.inmylanguage("Alternative Briefvorlage benutzen")
Label45.Caption = form1.inmylanguage("Strasse")
Label44.Caption = form1.inmylanguage("PF")
Label43.Caption = form1.inmylanguage("PLZP")
Label43.ToolTipText = form1.inmylanguage("Postleitzahl des Postfachs")
Label42.Caption = form1.inmylanguage("Ort")
Label41.Caption = form1.inmylanguage("PLZ")
Label40.Caption = form1.inmylanguage("Land")
Label46.Caption = form1.inmylanguage("Postf. geht vor")
Check3.ToolTipText = form1.inmylanguage("Postf. geht vor")
Label46.ToolTipText = transe("Haken entfernen, um nur Strasse/Ort zu benutzen")
Label31.Caption = form1.inmylanguage("PLZPostf")
Label18.Caption = form1.inmylanguage("(Postf.)")
Label17.Caption = form1.inmylanguage("Detailliste")
Label30.Caption = form1.inmylanguage("PLZ")
Label29.Caption = form1.inmylanguage("Pos.")
Label32.Caption = form1.inmylanguage("immer Bezeichnung anzeigen")
Label27.Caption = form1.inmylanguage("www")
Label27.ToolTipText = form1.inmylanguage("Internet-Adresse besuchen")
Image4.ToolTipText = form1.inmylanguage("Adresse löschen verboten")
Label24.Caption = form1.inmylanguage("Handy")
Label16.Caption = form1.inmylanguage("Handy")
Label10.Caption = form1.inmylanguage("Tel")
Label5.Caption = form1.inmylanguage("Tel")
Image1(0).ToolTipText = form1.inmylanguage("Kontakt löschen verboten")
Label39.Caption = form1.inmylanguage("Weitere Dokumente")
Label39.ToolTipText = form1.inmylanguage("Zusatzfelder je nach gewählten Kategorien")
suchen.ToolTipText = form1.inmylanguage("Adresse suchen")
Label38.Caption = form1.inmylanguage("Saison/Zeitraum:")
Label38.ToolTipText = form1.inmylanguage("Dieser Zeitraum interessiert mich")
Label37.Caption = form1.inmylanguage("In Kategorien:")
Label37.ToolTipText = form1.inmylanguage("Diese Personengruppen interessieren mich")
idshow.ToolTipText = form1.inmylanguage("Sortiername und eindeutiger Bezeichner der Adresse, Doppelklick zum umbenennen")
Label36.Caption = form1.inmylanguage("Termine:")
Label36.ToolTipText = form1.inmylanguage("Beteiligt an welchen Auftrittsterminen?")
Label35.Caption = form1.inmylanguage("zur.")
Label34.Caption = form1.inmylanguage("Zusatzfelder")
Label34.ToolTipText = form1.inmylanguage("Zusatzfelder je nach gewählten Kategorien")
Label33.Caption = form1.inmylanguage("Kontakte")
Label28.Caption = form1.inmylanguage("Fax")
Label26.Caption = form1.inmylanguage("formel")
Label25.Caption = form1.inmylanguage("Anzeigen:")
Kategorie.Caption = form1.inmylanguage("Kategorien:")
Label23.Caption = form1.inmylanguage("Land")
Label12.Caption = form1.inmylanguage("Bearbeiten:")
Label22.Caption = form1.inmylanguage("Name")
Label21.Caption = form1.inmylanguage("Schluss")
Label20.Caption = form1.inmylanguage("Anrede")
Label6.Caption = form1.inmylanguage("Fax")
Label19.Caption = form1.inmylanguage("Bundesland")
Label7.Caption = form1.inmylanguage("Pfad zur Mediendatei")
Label7.ToolTipText = form1.inmylanguage("Hier sind die Mediadateien gespeichert")
Label15.Caption = form1.inmylanguage("vor ")
Label14.Caption = form1.inmylanguage("Mx.Tg.")
Label13.Caption = form1.inmylanguage("Kd-Nr.")
Label11.Caption = form1.inmylanguage("www")
Label11.ToolTipText = form1.inmylanguage("Internet-Adresse besuchen")
Label9.Caption = form1.inmylanguage("geändert:")
Label8.Caption = form1.inmylanguage("Notizen")
Label4.Caption = form1.inmylanguage("Ort")
Label3.Caption = form1.inmylanguage("Strasse")
Label2.Caption = form1.inmylanguage("ID")
Label1.Caption = form1.inmylanguage("Name")
Label49.Caption = form1.inmylanguage("vertraulich")
Label49.ToolTipText = transe("Kennzeichnet die Adressedaten als vertraulich")
Label26.ToolTipText = transe("Abreden von") + " " + form1.anredeuser$
Label50.Caption = form1.inmylanguage("Tel2")
c4no = True:
Check4.value = 0: If form1.getusersetting("zusatzinfos", "normal") = "erweitert" Then Check4.value = 1
c4no = False
If Not form1.isfieldmissing("adresse", "optinternal") Then
  Label49.Visible = True
  intuse.Visible = True
End If
If Not form1.isfieldmissing("opt_allenummern", "vid") Then
  anumsel.Enabled = True
  knumsel.Enabled = True
  Command52.Enabled = True
  Command53.Enabled = True
End If
If form1.isfieldmissing("adresse", "opttel") Then
  opttel.text = "database needs extension."
  opttel.Enabled = False
  Label50.ToolTipText = "contact support to enable a second phonenumer."
Else
'  opttel.ToolTipText = form1.inmylanguage("Telefon")
  Label50.ToolTipText = form1.getusersetting("opttel", "Tel2")
  Label50.Caption = Label50.ToolTipText
End If
If form1.isfieldmissing("kontakt", "opttel") Then
  optktel.text = "database needs extension."
  optktel.Enabled = False
  Label51.ToolTipText = "contact support to enable a second phonenumer."
Else
  Label51.ToolTipText = form1.getusersetting("opttel", "Tel")
End If
Label50.Caption = form1.getusersetting("opttel", "Tel")
Label51.Caption = form1.getusersetting("optktel", "Tel")
adrnotz$ = form1.getusersetting("adressnotizdatei", "apinfo.rtf")
If adrnotz$ <> "" Then
  Command10.Width = 855
  kh2.Visible = True
  kh2ja.Visible = False
End If
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:gd1bez")
k1lrv% = Val(form1.mylastFormVar(Me.name, "usempth", "1"))
If k1lrv% <> 1 Then k1lrv% = 0
usempth.value = k1lrv%
k1lrv% = Val(form1.mylastFormVar(Me.name, "gd1bez", "0"))
If k1lrv% <> 1 Then k1lrv% = 0
gd1bez.value = k1lrv%

Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum"), 1400)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Typ"), 800)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Ort"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Honorar"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Bezeichnung / Hinweise"), 2000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Projekt"), 2000)
Set colHeader = gd1.ColumnHeaders.add(, , "AID", 0)

Set colHeader = gd2.ColumnHeaders.add(, , transe("Beschreibung"), 1200)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Ort"), 1000)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Strasse"), 1000)
Set colHeader = gd2.ColumnHeaders.add(, , "AID", 0)

Set colHeader = gd3.ColumnHeaders.add(, , transe("Beschreibung"), 1200)
Set colHeader = gd3.ColumnHeaders.add(, , transe("Ort"), 1000)
Set colHeader = gd3.ColumnHeaders.add(, , transe("Strasse"), 1000)
Set colHeader = gd3.ColumnHeaders.add(, , "AID", 0)

postanredea.AddItem transe("Frau")
postanredea.AddItem transe("Herr")
postanredea.AddItem transe("Herrn")
postanredea.AddItem transe("Firma")
postanredea.AddItem transe("Herrn und Frau")
postanredea.AddItem transe("Herr und Frau")
postanredea.AddItem transe("Familie")
postanredek.AddItem transe("Frau")
postanredek.AddItem transe("Herr")
postanredek.AddItem transe("Herrn")
postanredek.AddItem transe("Firma")
postanredek.AddItem transe("Herrn und Frau")
postanredek.AddItem transe("Herr und Frau")
postanredek.AddItem transe("Familie")
wd0 = 9995
wd0 = Shape7.Left + Shape7.Width + 160
'wd1 = 13020
wd1 = 12825
wd1 = Shape8.Left + Shape8.Width + 160
'hg0 = 6525
'hg1 = 8445
'hg0 = 6845
'hg1 = 8785
hg0 = 7215
hg1 = 9200
break% = 0
fl_rl3% = 0
nl3fl% = 0
nflds = 15
nfldsk = 9
nfldska = 5
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:fontsize")
s% = form1.myfontsize()
For i% = 0 To nflds: datf(i%).Font.Size = s%: Next i%
For i% = 0 To nfldsk: kdat(i%).Font.Size = s%: Next i%
For i% = 0 To nfldska: kadat(i%).Font.Size = s%: Next i%
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:vorlagen adressen_...")
tr = form1.vorlagendir() & "\adressen_*.rtf"
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:vorlagdir=" + form1.vorlagendir())
Combo1.Clear
Combo1.AddItem transe("keine Liste")
Combo1.AddItem "CSV-Export"
Combo1.AddItem "GEMA-Export"
If form1.cloud Then Combo1.AddItem "Cloud-Export"
If form1.kaldbok Then Combo1.AddItem transe("WebKalender")
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:tr=dir(" + tr + ")")
On Error Resume Next
tr = Dir(tr)
rrr = Err
On Error GoTo 0
If rrr <> 0 And form1.isstarting Then Call form1.startlog(form1.getuserid(), "Dir() failed, Fehler #" + trm(rrr))
If tr <> "" And rrr = 0 Then
  While tr <> ""
    If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:tr=" + tr)
    f$ = Mid(tr, InStr(tr, "_") + 1)
    f$ = Left(f$, InStr(f$, ".") - 1)
    Combo1.AddItem f$
    tr = Dir
  Wend
End If
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:position ermitteln")
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Top = form1.Top + form1.Height + 40
End If
Call form1.formpos(Me)

If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:formular leeren")
Call nulldsp
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:briefvorlagen finden")
altbvorl.text = form1.meinebriefvorlage()
Command3.ToolTipText = transe("Brief schreiben") + ", " + transe("Vorlage") + ": " & altbvorl.text
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:connect-string ermitteln")
dbpara$ = form1.getconnstr()
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:ado-connect-string ermitteln")
adopara$ = form1.getadoconnstr()
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:von-bis für termine")
toffsvon = Val(form1.mylastFormVar(Me.name, "toffsvon", "0"))
toffsbis = Val(form1.mylastFormVar(Me.name, "toffsbis", "0"))
Text1.text = toffsvon
Text2.text = toffsbis

If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:höhe und breite setzen")
Me.Width = wd0
Me.Height = hg0
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:Mozillaprofile")
cf$ = form1.getusersetting("Mozillaprofile") & "\Calendar\CalendarManager.rdf"
If exist(cf$) = 1 Then
  icalconf.Enabled = True
Else
  cf$ = form1.getusersetting("iCalendarprofile")
  If exist(cf$) = 1 Then icalconf.Enabled = True
End If
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:formular anzeigen")
If form1.getusersetting("adressenumbenennen", "verboten") <> "erlaubt" Then adr_ren.Enabled = False
autokategorie$ = form1.getusersetting("automarkierekategorie", "")
usekpos = True
If form1.isfieldmissing("kontakt", "opt_kpos") Then usekpos = False
If form1.usemenu <> "ja" Then
  adr.Visible = False
  knt.Visible = False
  hlp.Visible = False
  dok.Visible = False
  kat.Visible = False
  hg0 = 7260 - 240
  Me.Height = hg0
  hg1 = 9200 - 240
End If
knt.Enabled = False
mytopmerk = Me.Top
On Error Resume Next
'If form1.isfieldmissing("opt_adresspool", "id") Then
  cadrpbez.Visible = False
  ckadrpbez.Visible = False
'End If
Show
If Me.Left = 20 And Me.Top = 20 Then
  Me.Left = form1.Left
  form1.Width = Me.Width
End If
On Error GoTo 0
DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:show done")
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:shape4")
Shape4.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:shape2")
Shape2.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:shape3")
Shape3.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:shape1")
Shape1.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:shape7")
Shape7.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
If form1.isstarting Then Call form1.startlog(form1.getuserid(), "shwAdrDetail:Load done")

End Sub
Public Sub refreshadrdetail(vid$, contact$)
Dim rtmp As ADODB.Recordset, pr As ADODB.Recordset, prv2$, r As ADODB.Recordset, c$, prvid$
Dim stmp As ADODB.Recordset, i%, wert$, na$, tpid$, renum As Boolean, updsv As Boolean
Dim s As ADODB.Recordset, rrr, tz$, sflds As Integer, cmd$, tabkz As String

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "refreshadrdetail"
Call form1.dbg2f("shwAdrDetail:refreshadrdetail")
Call form1.setAuftrittsdruckFuerAdresse("")
srchit% = 0
break% = 0
Call nulldsp
cadrpbez.text = "Standard"
ckadrpbez.text = "Standard"
datf(0).text = vid$
kh2.Visible = True
kh2ja.Visible = False
form1.adrmerkid = datf(0).text
Label7.Caption = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(vid$)
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM adresse where id='" & vid$ & "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not rtmp.EOF Then
  rtmp.MoveFirst
  nodbupd = True
  intuse.value = 0
  nodbupd = False
  Command34.Enabled = True
  Command31.Enabled = True
  If intuse.Visible Then
    i% = Val(trm0(rtmp!optinternal))
    If i% > 0 Then
      nodbupd = True
      intuse.value = 1
      nodbupd = False
      Command34.Enabled = False
      Command31.Enabled = False
    End If
  End If
  For i% = 0 To nflds
    On Error Resume Next
    wert$ = trm(rtmp.Fields(i%).value)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 And wert$ <> "" And i% <> 16 Then
      datf(i%).text = rtmp.Fields(i%)
      If i% = 14 Then
        Set stmp = New ADODB.Recordset
        stmp.CursorLocation = adUseServer
        c$ = "select wert from sysvars where owner='sysvar_system_landeskennung_" & datf(14).text & "'"
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not stmp.EOF Then
          If stmp!wert <> datf(14).text Then datf(14).text = stmp!wert
        End If
      End If
    End If
  Next i%
  If id$ <> "" Then Call form1.chkallnums(id$, "-1", "email", datf(6).text)
  plzp.text = trm(rtmp!plzpostfach)
  postf.text = trm(rtmp!postfach)
  postanredea.text = trm(rtmp!postanrede)
  If Not form1.isfieldmissing("adresse", "opttel") Then
    opttel.text = trm(rtmp!opttel)
  Else
    opttel.text = "database needs extension."
  End If
End If
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
If Not nexist(dn$ + "\" + adrnotz$) Then
  kh2.Visible = False
  kh2ja.Visible = True
End If
c$ = "select wert from sysvars where owner='sysvar_" + form1.getuserid() + "_zzzadr_sticky_" + trm(idshow.Caption) + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
stcky_igno = True
stcky.value = 0
If rrr = 0 Then
  If Not rtmp.EOF Then stcky.value = 1
End If
stcky_igno = False
tz$ = form1.getusersetting("plzort-" & trm(datf(14)), "L P O")
Label4.ToolTipText = tz$ & " (Doppelklick ändert Land-PLZ-Ort Reihenfolge)"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,name,position"
If usekpos Then c$ = c$ + ",opt_kpos"
c$ = c$ + " FROM kontakt where vid ='" + vid$ + "'"
If usekpos Then c$ = c$ + " order by opt_kpos"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

renum = False
If Not rtmp.EOF Then
rtmp.MoveFirst
renum = False
While Not rtmp.EOF
  na$ = ""
  If Not IsNull(rtmp!name) Then
    na$ = rtmp!name
    If trm(rtmp!Position) <> "" Then na$ = na$ + " (" + rtmp!Position + ")"
    If usekpos Then
      If trm(rtmp!opt_kpos) = "" Then renum = True
    End If
  End If
  klist.AddItem form1.crlffake(na$)
  idxlist.AddItem rtmp!id
  rtmp.MoveNext
Wend
End If
If renum Then
  form1.noalarms = True
  For i% = 0 To idxlist.ListCount - 1
    c$ = "update kontakt set opt_kpos=" + trm(i%) + " where id='" + idxlist.List(i%) + "'"
    Call form1.sqlqry(c$)
  Next i%
  form1.noalarms = False
End If
Call kposbuttonset
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,typ,wert FROM adresstyp where vid ='" + vid$ + "' and (kid='-1' or isnull(kid)) order by wert"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

prv2$ = ""
prvid$ = ""
While Not rtmp.EOF
  na$ = "": If Not IsNull(rtmp!typ) Then na$ = transe(rtmp!typ)
  If Not IsNull(rtmp!wert) Then na$ = na$ + ": " + rtmp!wert
  If Left(na$, 4) <> "rel:" Then

  If na$ = prv2$ Then
    c$ = "SELECT * From auftritthigru where auftrittsid='" & datf(0).text & "." & rtmp!id & "' and feldname='" + rtmp!typ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF And break% = 0
      Debug.Print datf(0).text & "." & rtmp!id & " ==> " & datf(0).text & "." & prvid$
      r.MoveNext
    Wend
    Debug.Print "löschen:" & datf(0).text & "." & rtmp!id & "' and feldname='" + rtmp!typ + "'"
    c$ = "delete from adresstyp where id='" + rtmp!id + "'"
    Call form1.sqlqry(c$)
  Else
    prv2$ = na$
    prvid$ = datf(0).text & "." & rtmp!id
    List1.AddItem transe(na$)
  End If

  Else
    tabkz = form1.getusersetting("relabkz_" + Mid(trm(rtmp!typ), 5), Mid(trm(rtmp!typ), 5))
    List1b.AddItem tabkz + ":" + trm(rtmp!wert)
  End If
  rtmp.MoveNext
Wend

sflds = 0
For i% = sflds To nfldsk
  kfname(i% - sflds).Caption = ""
  kdat(i% - sflds).text = ""
Next i%
For i% = 0 To nfldska
  kadat(i%).text = ""
Next i%
Anrede.text = ""
Abrede.text = ""
If klist.ListIndex < 0 And trm(datf(0).text) <> "" Then
  Anrede.text = form1.meineanrede("-1." & datf(0).text)
  Abrede.text = form1.meineabrede("-1." & datf(0).text)
End If
rtmp.Close
If Not form1.isfieldmissing("opt_prios", "id") Then
    prio.Enabled = True
    prio.Visible = True
    If vid$ <> "" Then
      cmd$ = "SELECT * FROM opt_prios where evnt= 'A:" & vid$ & "' and userid='" + form1.getuserid() + "'"
      Set pr = New ADODB.Recordset
      pr.CursorLocation = adUseServer
      Call form1.dbg2f("auftritt.showrec:" & cmd$)
      On Error Resume Next
      rrr = form1.adoopen(pr, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        prio.text = ""
      End If
      If Not pr.EOF Then
        prio.text = trmx1(pr!prio)
      Else
        prio.text = ""
      End If
    End If
Else
    prio.Enabled = False
    prio.Visible = False
End If
Label6.Visible = True
Call mlist
Call rlist4
srchit% = 1
datf(13).ToolTipText = transe("Postleitzahl")
datf(3).ToolTipText = transe("Ort")
If Command14.Caption = transe("ohne Zusätze") Then Call Command15_Click
Command4.Enabled = False
BackColor = form1.cleancolor()
Command5.Enabled = False
knt_sav.Enabled = False
Command6.Visible = False
Command12.Enabled = False
knt.Enabled = False
If klist.ListCount > 0 Then knt.Enabled = True
DoEvents
If vid$ <> "" And Not form1.isfieldmissing("opt_adresspool", "id") Then
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  c$ = "select id,Beschreibung from opt_adresspool where vid='" + vid$ + "' and kid='-1'"
  r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
  updsv = False
  If r.EOF Then
    updsv = True
  Else
    updsv = True
    While Not r.EOF And updsv
      If LCase(trm(r!Beschreibung)) = "standard" Then updsv = False
      r.MoveNext
    Wend
  End If
  If updsv Then Call svadr("Standard")
End If

If Not form1.isfieldmissing("opt_repertoire", "id") Then
  If contact$ = "" Or contact$ = "-1" Then
    c$ = vid$
  Else
    c$ = contact$ + "{" + vid$ + "}"
  End If
  If form1.isoftype(c$, "Künstler") <> "-1" Or form1.isoftype(c$, "Dirigent") <> "-1" Or form1.isoftype(c$, "Orchester") <> "-1" Then
    repert.Visible = True
  End If
End If
If contact$ <> "" Then
  For i% = 0 To klist.ListCount - 1
    'If InStr(klist.List(i%), contact$) = 1 Or InStr(contact$, klist.List(i%)) = 1 Or idxlist.List(i%) = contact$ Then
    If klist.List(i%) = contact$ Or idxlist.List(i%) = contact$ Then
      klist.ListIndex = i%
      BackColor = form1.cleancolor()
      Exit Sub
    End If
  Next i%
End If
Command12.Enabled = True
'
'BEWARE EXIT SUB ABOVE
'
'ok 4 this:
If trm(datf(0).text) <> "" Then
  If Anrede.text = "" Then Anrede.text = form1.getusersetting("StandardAnrede", "")
  If Abrede.text = "" Then Abrede.text = form1.getusersetting("StandardAbrede", "")
  c$ = "select * from dochist where adresse='" + trm(datf(0).text) + "' and kontakt='-1' and doctyp='" + transe("Datenänderung") + "';"
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    Label9.Caption = transe("geändert:")
    If Not s.EOF Then
      Label9.Caption = Label9.Caption + " von " + trm(s!Owner)
      datf(8).text = datfromsql(word1(trm(s!erstellt))) + " " + word2bis(trm(s!erstellt))
    End If
  End If
  Me.BackColor = form1.cleancolor()
  If intuse.value <> 0 Then Me.BackColor = Val(form1.getusersetting("internalcolor", "12828927"))
  knt_sav.Enabled = False
  adr_sav.Enabled = False
End If
DoEvents
If autokategorie$ <> "" And List1.Visible = True Then
  For i% = 0 To List1.ListCount - 1
    If InStr(List1.List(i%), autokategorie$) = 1 Then
      List1.ListIndex = i%
      Exit For
    End If
  Next i%
End If
End Sub

Sub nulldsp()
Dim i%, sanz As Long, ss%, sd%, rrr, sey%, d0, se%, sed, l$

'd2infile = "shwAdrDetail": d2insub = "nulldsp"
If List1b.Visible Then Call Command42_Click
Unload bezlist
l1bdont = False
currentk = "-1"
gd2.ListItems.Clear
gd2.Visible = False
gd3.ListItems.Clear
gd3.Visible = False
For i% = 0 To nflds: datf(i%) = "": Next i%
For i% = 0 To nfldsk: kdat(i%) = "": kdat(i%).Enabled = False: Next i%
For i% = 0 To nfldska: kadat(i%) = "": kadat(i%).Enabled = False: Next i%
postanredea.text = ""
postanredek.text = ""
klist.Clear
idxlist.Clear
opttel.text = ""
List1.Clear
List1b.Clear
knt_sav.Enabled = False
adr_sav.Enabled = False
Command4.Enabled = False
adr_sav.Enabled = False
BackColor = form1.cleancolor()
p1offs% = 0
Combo1.text = transe("keine Liste")
Command5.Enabled = False
knt_sav.Enabled = False
Command6.Visible = False
Command13.Visible = False
Check1.value = 1
Check2.value = 1
season.Clear
season.Enabled = False
ss% = Val("0" & trm(form1.getusersetting("saisonstart")))
sd% = Val("0" & trm(form1.getusersetting("saisondauer")))
If sd% = 0 Then sd% = 12
If ss% <> 0 Then
  sanz = 4
  On Error Resume Next
  sanz = Val(form1.getusersetting("adressenzeigejahreinsaisonwahl", "0"))
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then sanz = 4
  season.Enabled = True
  sey% = apyear(Date) - sanz / 2
  If sanz = 0 Then
    sey% = apyear(Date) - 10
    sanz = 10
  End If
  While sey% < apyear(Date) + sanz / 2
    d0 = CDate("1." & trm(ss%) & "." & trm(sey%))
    se% = (ss% + sd%)
    If se% > 12 Then
      se% = (se% - 1) Mod 12 + 1
      sey% = sey% + 1
    End If
    ss% = se%
    sed = CDate("1." & trm(se%) & "." & trm(sey%)) - 1
    l$ = trm(d0) & "-" & trm(sed)
    season.AddItem l$
  Wend
End If
If form1.isfieldmissing("adresse", "opttel") Then
  opttel.text = "database needs extension."
  opttel.Enabled = False
  Label50.ToolTipText = "contact support to enable a second phonenumer."
Else
  opttel.text = ""
End If

If form1.isfieldmissing("kontakt", "opttel") Then
  optktel.text = "database needs extension."
  optktel.Enabled = False
  Label51.ToolTipText = "contact support to enable a second phonenumer."
Else
  optktel.text = ""
End If

List4.Clear
List2.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "shwAdrDetail": d2insub = "Form_Unload"
break% = 1
Call savecheck
form1.adrmerkid = ""
Call mynormsize
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub
Public Sub mynormsize()
If Command14.Caption = transe("ohne Zusätze") Then Call Command14_Click
If Command18.Caption = transe("ohne Medien") Then Call Command18_Click

End Sub
Private Sub gd1_DblClick()
Dim id$, lvitem

'd2infile = "shwAdrDetail": d2insub = "gd1_DblClick"
Set lvitem = gd1.SelectedItem
id$ = ""
On Error Resume Next
id$ = lvitem.SubItems(6)
On Error GoTo 0
If id$ = "" Then Exit Sub
Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(id$, 0)

End Sub

Private Sub gd1bez_Click()
'd2infile = "shwAdrDetail": d2insub = "gd1bez_Click"
Call form1.setmylastFormVar(Me.name, "gd1bez", trm(gd1bez.value))
Call rlist2
Label36.Caption = form1.inmylanguage("Termine: ") + trm(List2.ListCount)
DoEvents
Call rcombo2
End Sub

Private Sub gd1show_Click()
Dim i%

'd2infile = "shwAdrDetail": d2insub = "gd1show_Click"
If gd1show.value = 1 Then
  gd1.Visible = True
  Anrede.Visible = False
  Abrede.Visible = False
  postanredek.Visible = False
  Command17(1).Visible = False
  Command39.Visible = False
  Command36.Visible = False
  List4.Visible = False
  gd1bez.Visible = True
  List2.Visible = False
  Label20.Visible = False
  Label32.Visible = True
  For i% = 2 To 9
    kdat(i%).Visible = False
  Next i%
  For i% = 0 To nfldska
    kadat(i%).Visible = False
  Next i%
Else
  gd1.Visible = False
  Anrede.Visible = True
  Abrede.Visible = True
  postanredek.Visible = True
  gd1bez.Visible = False
  Label32.Visible = False
  Label20.Visible = True
  Command17(1).Visible = True
  Command39.Visible = True
  Command36.Visible = True
  List4.Visible = True
  List2.Visible = True
  For i% = 2 To 9
    kdat(i%).Visible = True
  Next i%
  For i% = 0 To nfldska
    kadat(i%).Visible = True
  Next i%
End If
If gd1upd% = 1 Then Call form1.setmylastFormVar(Me.name, "gd1show", trm(gd1show.value))

End Sub

Private Sub gd2_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub gd3_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub gd2_DblClick()
Dim c$, r As ADODB.Recordset, cid$
Dim id$, lvitem, crkont As Boolean

gd2.Visible = False
cid$ = datf(0).text
If cid$ = "" Then Exit Sub

crkont = CtrlKey()
'd2infile = "shwAdrDetail": d2insub = "gd1_DblClick"
Set lvitem = gd2.SelectedItem
id$ = ""
On Error Resume Next
id$ = lvitem.SubItems(3)
On Error GoTo 0
If id$ = "" Then Exit Sub
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c$ = "select * from opt_adresspool where id='" + id$ + "'"
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  If Not crkont Then
    datf(3).text = trm(r!ort)
    datf(2).text = trm(r!strasse)
    datf(13).text = trm(r!plz)
    postf.text = trm(r!postfach)
    datf(14).text = trm(r!land)
    plzp.text = trm(r!plzpostfach)
    cadrpbez.text = trm(r!Beschreibung)
    If cadrpbez.text = "Standard" Then
      cadrpbez.Visible = False
    Else
      cadrpbez.Visible = True
    End If
  Else
    Call Command7_Click: DoEvents
    kdat(2).text = trm(r!Beschreibung)
    kadat(3).text = trm(r!ort)
    kadat(0).text = trm(r!strasse)
    kadat(2).text = trm(r!plz)
    kadat(5).text = trm(r!postfach)
    kadat(1).text = trm(r!land)
    kadat(4).text = trm(r!plzpostfach)
  End If
  Me.BackColor = form1.cleancolor()
End If
End Sub

Private Sub gd2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c$, r As ADODB.Recordset, cid$
Dim id$, lvitem

cid$ = datf(0).text
If cid$ = "" Then Exit Sub


If KeyCode = 46 Or KeyCode = 8 Then
  Set lvitem = gd2.SelectedItem
  id$ = ""
  On Error Resume Next
  id$ = lvitem.SubItems(3)
  On Error GoTo 0
  If id$ = "" Then Exit Sub
  c$ = "delete from opt_adresspool where id='" + id$ + "'"
  Call form1.sqlqry(c$)
  gd2.Visible = False
End If

End Sub

Private Sub gd3_DblClick()
Dim c$, r As ADODB.Recordset, cid$
Dim id$, lvitem

gd3.Visible = False
cid$ = datf(0).text
If cid$ = "" Then Exit Sub

'd2infile = "shwAdrDetail": d2insub = "gd1_DblClick"
Set lvitem = gd3.SelectedItem
id$ = ""
On Error Resume Next
id$ = lvitem.SubItems(3)
On Error GoTo 0
If id$ = "" Then Exit Sub
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c$ = "select * from opt_adresspool where id='" + id$ + "'"
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  kadat(3).text = trm(r!ort)
  kadat(0).text = trm(r!strasse)
  kadat(2).text = trm(r!plz)
  kadat(5).text = trm(r!postfach)
  kadat(1).text = trm(r!land)
  kadat(4).text = trm(r!plzpostfach)
  ckadrpbez.text = trm(r!Beschreibung)
  If ckadrpbez.text = "Standard" Then
    ckadrpbez.Visible = False
  Else
    ckadrpbez.Visible = True
  End If
  Me.BackColor = form1.cleancolor()
End If

End Sub

Private Sub gd3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c$, r As ADODB.Recordset, cid$
Dim id$, lvitem

cid$ = datf(0).text
If cid$ = "" Then Exit Sub


If KeyCode = 46 Or KeyCode = 8 Then
  Set lvitem = gd3.SelectedItem
  id$ = ""
  On Error Resume Next
  id$ = lvitem.SubItems(3)
  On Error GoTo 0
  If id$ = "" Then Exit Sub
  c$ = "delete from opt_adresspool where id='" + id$ + "'"
  Call form1.sqlqry(c$)
  gd3.Visible = False
End If

End Sub

Private Sub higrusuch_Click()

Load higruselect2
On Error Resume Next
Call higruselect2.SetFocus
On Error Resume Next
Exit Sub

' hmmm, .... needed?
If List1.ListIndex >= 0 Then
  Load higruselect2
  DoEvents
  Exit Sub
End If
Load higruselect
End Sub

Private Sub hlp_hlp_Click()
Call Command28_Click
End Sub

Private Sub idshow_Change()
Unload zusinf
Unload bezlist
Unload dochist2
End Sub

Private Sub idshow_DblClick()
Dim r As ADODB.Recordset, p%, l$, c$, fnam$, lcap$, altwert$, neuwert$, cmd$
Dim ask As Integer, neuid$, rrr, ttt$, altid

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "idshow_DblClick"
If trm(idshow.Caption) = "" Then Exit Sub
If form1.getusersetting("adressenumbenennen", "verboten") <> "erlaubt" Then
  Clipboard.Clear
  Clipboard.settext idshow.Caption
  MsgBox "AddressID: " + idshow.Caption + vbCrLf + "was copied to clipboard."
  Exit Sub
End If
  
  ask = MsgBox(transe("Das Umbenennen einer Adresse ist potentiell gefährlich.") & vbCrLf & transe("z.B. werden die zugehörigen Verzeichnisse nicht umbenannt.") & vbCrLf & transe("Haben Sie eine Datensicherung durchgeführt?."), vbYesNo + vbCritical + vbDefaultButton2, transe("Sortiernamen ändern?"))
  If ask <> vbYes Then Exit Sub
  neuid$ = trm(InputBox("Sortiernamen ändern", idshow.Caption, idshow.Caption))
  neuid$ = strrepl(neuid$, "/", "_")
  neuid$ = strrepl(neuid$, "'", "´")
  neuid$ = strrepl(neuid$, "&", "_")
  If neuid$ = "" Then Exit Sub
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  c$ = "SELECT * FROM adresse where id='" & neuid$ & "';"
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    MsgBox transe("Dieser Sortiername existiert bereits.")
    Exit Sub
  End If
  MousePointer = 11: DoEvents
  altid = idshow.Caption
  form1.sqlqry ("update adresse set id='" & neuid$ & "' where id='" & altid & "'")
  If Not form1.isfieldmissing("opt_topics", "id") Then
    c$ = "select id,owner,wert from sysvars where owner like 'sysvar_system_tlnk_%_" + altid + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
      l$ = Left(trm(r!Owner), Len(trm(r!Owner)) - Len(altid)) + neuid$
'Debug.Print r!Owner; " ("; r!wert; ")"; vbCrLf; l$
      c$ = "update sysvars set owner='" + l$ + "' where id='" + trm(r!id) + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
    Call form1.sqlqry("update opt_topics set vid='" & neuid$ & "' where vid='" & altid & "'")
    Call form1.sqlqry("update sysvars set wert='" & neuid$ & "' where wert='" & altid & "'")
  End If
  If Not form1.isfieldmissing("opt_allenummern", "vid") Then form1.sqlqry ("update opt_allenummern set vid='" & neuid$ & "' where vid='" & altid & "'")
  If Not form1.isfieldmissing("opt_adresspool", "id") Then
    form1.sqlqry ("update opt_adresspool set vid='" & neuid$ & "' where vid='" & altid & "'")
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    c$ = "SELECT id FROM opt_adresspool where id like '" + altid + "%';"
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
'Debug.Print r!id;
      l$ = neuid$ + Mid(r!id, Len(altid) + 1)
      c$ = "update opt_adresspool set id='" + l$ + "' where id='" + r!id + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
  End If
  form1.sqlqry ("update adressgruppen set adressid='" & neuid$ & "' where adressid='" & altid & "'")
  form1.sqlqry ("update adresstyp set vid='" & neuid$ & "' where vid='" & altid & "'")
  form1.sqlqry ("update adresstyp set wert='" & neuid$ & "' where wert='" & altid & "'")
  form1.sqlqry ("update anreden set kid='-1." & neuid$ & "' where kid='-1." & altid & "'")
  If Not form1.isfieldmissing("opt_prios", "id") Then
    form1.sqlqry ("update opt_prios set evnt='" & neuid$ & "' where evnt='" & altid & "'")
  End If
  If Not form1.isfieldmissing("opt_numbers", "id") Then
    form1.sqlqry ("update opt_numbers set vid='" & neuid$ & "' where vid='" & altid & "'")
  End If
  If Not form1.isfieldmissing("opt_repertoire", "id") Then
    form1.sqlqry ("update opt_repertoire set vid='" & neuid$ & "' where vid='" & altid & "'")
  End If
  lcap$ = LCase(altid)
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  c$ = "SELECT id,wert FROM adresstyp where ((InStr(wert,'{" & altid & "}'))>0)"
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    p% = InStr(r!wert, "{")
    If p% > 1 Then
      ttt$ = trm(Left(r!wert, p% - 1)) + " {" + neuid$ + "}"
      c$ = "update adresstyp set wert='" + ttt$ + "' where id='" + r!id + "'"
      form1.sqlqry (c$)
    End If
    r.MoveNext
  Wend
      
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    c$ = "SELECT id,auftrittsid FROM auftritthigru where auftrittsid like '" + altid + "%';"
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
      l$ = neuid$ + Mid(r!auftrittsid, Len(altid) + 1)
'Debug.Print r!auftrittsid; " - "; r!id; " - "; l$
      c$ = "update auftritthigru set auftrittsid='" + l$ + "' where id='" + r!id + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend

  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  c$ = "SELECT * FROM auftritthigru where (((InStr(FeldDaten,'" & altid & "'))=1));"
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    l$ = trm(r!felddaten)
    If LCase(l$) = lcap$ Then
      form1.sqlqry ("update auftritthigru set felddaten='" & neuid$ & "' where id='" & r!id & "'")
      l$ = utabn(trm(r!auftrittstyp))
      If l$ <> "" Then
        fnam$ = trm(r!feldname)
        cmd$ = "update usr_" & l$ & " set " & fnam$ & "='" & neuid$ & "' where id='" & r!auftrittsid & "'"
        Call form1.sqlqry(cmd$)
        DoEvents
      End If
    End If
    r.MoveNext
  Wend
  r.Close
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  c$ = "SELECT * FROM auftritthigru where (((InStr(FeldDaten,'{" & altid & "}'))>0));"
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    l$ = trm(r!felddaten): altwert$ = l$
Debug.Print l$
    p% = InStr(l$, "{")
    If LCase(Mid$(l$, p%)) = "{" & lcap$ & "}" Then
      cmd$ = "update auftritthigru set felddaten='" & Left(altwert$, p%) & neuid$ & "}' where id='" & r!id & "'"
      Call form1.sqlqry(cmd$)
      l$ = utabn(trm(r!auftrittstyp))
      fnam$ = trm(r!feldname)
      cmd$ = "update usr_" & l$ & " set " & fnam$ & "='" & Left(altwert$, p%) & neuid$ & "}' where id='" & r!auftrittsid & "'"
      Call form1.sqlqry(cmd$)
    End If
    r.MoveNext
  Wend
  r.Close
  altwert$ = altid & " [Wiedervorlage] Adresse:" & altid
  neuwert$ = neuid$ & " [Wiedervorlage] Adresse:" & neuid$
  form1.sqlqry ("update todolist set Betreff='" & neuwert$ & "' where Betreff='" & altwert$ & "'")
    
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    c$ = "SELECT id,Betreff FROM todolist where Betreff like '%" + altid + "%';"
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
Debug.Print r!betreff; " - "; r!id
      p% = InStr(r!betreff, altid)
      l$ = Left(r!betreff, p%)
      neuwert$ = l$ + neuid$ + Mid(r!betreff, Len(altid) + 1)
      c$ = "update todolist set Betreff='" + neuwert$ + "' where id='" + r!id + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    c$ = "SELECT id,Betreff FROM todolist where Betreff like '%" + altid + "';"
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
Debug.Print r!betreff; " - "; r!id
      p% = InStr(r!betreff, altid)
      l$ = Left(r!betreff, p%)
      neuwert$ = l$ + neuid$
      c$ = "update todolist set Betreff='" + neuwert$ + "' where id='" + r!id + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
 
  form1.sqlqry ("update bestellung set bestelleradresse='" + neuid$ & "' where bestelleradresse='" & altid & "'")
  form1.sqlqry ("update bestellung set lieferadresse='" + neuid$ & "' where lieferadresse='" & altid & "'")
  form1.sqlqry ("update bplan set adressid='" + neuid$ & "' where adressid='" & altid & "'")
  form1.sqlqry ("update dochist set adresse='" + neuid$ & "' where adresse='" & altid & "'")
  form1.sqlqry ("update finanzen set an='" + neuid$ & "' where an='" & altid & "'")
  form1.sqlqry ("update finanzen set von='" + neuid$ & "' where von='" & altid & "'")
  form1.sqlqry ("update hbabotermine set adrid='" + neuid$ & "' where adrid='" & altid & "'")
  form1.sqlqry ("update hblist set hid='" + neuid$ & "' where hid='" & altid & "'")
  form1.sqlqry ("update hbplist set hid='" + neuid$ & "' where hid='" & altid & "'")
  form1.sqlqry ("update kassenbuch set vonid='" + neuid$ & "' where vonid='" & altid & "'")
  form1.sqlqry ("update kontakt set vid='" + neuid$ & "' where vid='" & altid & "'")
  form1.sqlqry ("update taliste set orchester='" + neuid$ & "' where orchester='" & altid & "'")
  form1.sqlqry ("update taliste set dirigent='" + neuid$ & "' where dirigent='" & altid & "'")
  form1.sqlqry ("update tplan set dirigent='" + neuid$ & "' where dirigent='" & altid & "'")
  form1.sqlqry ("update tplan set orchester='" + neuid$ & "' where orchester='" & altid & "'")
  form1.sqlqry ("update tplan set veranstalter='" + neuid$ & "' where veranstalter='" & altid & "'")
  form1.sqlqry ("update tplan set solist='" + neuid$ & "' where solist='" & altid & "'")
  form1.sqlqry ("update tplan set Projektbetreuer='" + neuid$ & "' where Projektbetreuer='" & altid & "'")
  form1.sqlqry ("update tpwernoch set kid='" + neuid$ & "' where kid='" & altid & "'")
  form1.sqlqry ("update auftritthigru set auftrittsid='" + neuid$ & "' where auftrittsid='" & altid & "'")
  idshow.Caption = neuid$
  datf(0).text = neuid$
  MousePointer = 0
End Sub

Private Sub Image1_Click(Index As Integer)
Dim fn$, neuid As String, cid$, kid$

'd2infile = "shwAdrDetail": d2insub = "Image1_Click"

If knt_sav.Enabled Then Call savecheck
If Index = 19 Then
  If form1.isfieldmissing("opt_adresspool", "id") Then
    MsgBox ("Die erforderliche Tabelle opt_adresspool fehlt." + vbCrLf + "Bitte kontaktieren Sie den Support.")
    Exit Sub
  End If
  cid$ = datf(0).text
  If cid$ = "" Then Exit Sub
  neuid = trm(InputBox(transe("Adresse merken als"), transe("Diese Adresse merken")))
  If trm(neuid) = "" Then Exit Sub
  Call svadr(neuid)
  Me.BackColor = form1.cleancolor()
End If
If Index = 1 Then
  kid$ = kdat(0).text
  If kid$ = "" Then Exit Sub
  If form1.isfieldmissing("opt_adresspool", "id") Then
    MsgBox ("Die erforderliche Tabelle opt_adresspool fehlt." + vbCrLf + "Bitte kontaktieren Sie den Support.")
    Exit Sub
  End If
  cid$ = datf(0).text
  If cid$ = "" Then Exit Sub
  neuid = trm(InputBox(transe("Adresse merken als"), transe("Diese Adresse merken")))
  If trm(neuid) = "" Then Exit Sub
  Call svkadr(neuid)
  Me.BackColor = form1.cleancolor()
End If
End Sub

Private Sub inclcont_Click()
Dim w$
w$ = "nein": If inclcont.value = 1 Then w$ = "ja"
Call form1.setusersetting("terminlisteauchadresse", w$)
End Sub

Private Sub intuse_Click()
Dim c$

If intuse.value = 0 Then
  c$ = "0"
  Me.BackColor = form1.cleancolor()
  Command34.Enabled = True
  Command31.Enabled = True
Else
  c$ = "1"
  Me.BackColor = Val(form1.getusersetting("internalcolor", "12828927"))
  Command34.Enabled = False
  Command31.Enabled = False
End If
If Not nodbupd Then
  c$ = "update adresse set optinternal=" + c$ + " where id='" + datf(0).text + "'"
  Call form1.sqlqry(c$)
End If
End Sub

Private Sub kadat_Change(Index As Integer)
'd2infile = "shwAdrDetail": d2insub = "kadat_Change"
Command5.Enabled = True
knt_sav.Enabled = True
BackColor = form1.dirtycolor()
End Sub

Private Sub kadat_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'd2infile = "shwAdrDetail": d2insub = "kadat_KeyDown"
If Index < 4 Then Call esckcntinc(KeyCode)

End Sub

Private Sub kat_add_Click()
Call Command8_Click
End Sub

Private Sub kat_del_Click()
Call Command9_Click
End Sub

Private Sub kdat_Change(Index As Integer)
'd2infile = "shwAdrDetail": d2insub = "kdat_Change"
Command5.Enabled = True
knt_sav.Enabled = True
BackColor = form1.dirtycolor()
'telfaxhandy
If Index = 4 Or Index = 3 Or Index = 6 Then
kdat(7).text = onlynums(kdat(4).text) + " " + onlynums(kdat(3).text) + " " + onlynums(kdat(6).text)
End If

End Sub


Private Sub kdat_DblClick(Index As Integer)
Dim w$, z$, i%, f$

'd2infile = "shwAdrDetail": d2insub = "kdat_DblClick"
If Index = 3 Or Index = 4 Or Index = 6 Then

w$ = kdat(Index).text
For i% = 1 To Len(w$)
  z$ = Mid$(w$, i%, 1)
  If isdigit(z$) > 0 Then
    f$ = form1.wavdir() & "\" & z$ & ".wav"
    If exist(f$) <> 0 Then
      Call sndPlaySound(f$, SND_SYNC)
    End If
  End If
  DoEvents
Next i%

End If

End Sub

Private Sub kdat_LostFocus(Index As Integer)
kdat(Index).text = strrepl(kdat(Index).text, "'", "´")
End Sub

Private Sub kh2ja_Click()
Call kh2_Click
End Sub
Private Sub kh2_Click()
Dim dn$

On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error GoTo 0
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
If nexist(dn$ + "\" + adrnotz$) Then
  On Error Resume Next
  Call FileCopy(form1.vorlagendir() + "\" + adrnotz$, dn$ + "\" + adrnotz$)
  On Error GoTo 0
End If
If nexist(dn$ + "\" + adrnotz$) Then
  MsgBox (transe("Die Datei") + " " + adrnotz$ + " " + transe("kann nicht gefunden werden.") + vbCrLf + "Bitte stellen Sie sicher, dass diese sich im Vorlagenverzeichnis befindet.")
  Exit Sub
End If
On Error Resume Next
Call form1.openthisdoc(dn$ + "\" + adrnotz$, "")
On Error GoTo 0

End Sub

Private Sub klist_Click()
Dim rtmp As ADODB.Recordset, c$, prvid$, rrr, dn$, tabkz As String
Dim s As ADODB.Recordset, sflds, i%, bg, klistid, tr As String

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "klist_Click"
Call savecheck
Unload zusinf
Unload bezlist
sflds = 0
i% = 0
List1b.Clear
bg = BackColor
For i% = sflds To nfldsk
  kdat(i%).Enabled = True
  kfname(i% - sflds).Caption = form1.sqla.TableDefs("kontakt").Fields(i%).name
  kdat(i% - sflds).text = ""
Next i%
For i% = 0 To nfldska
  kadat(i%).Enabled = True
  kadat(i%).text = ""
Next i%
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM kontakt where id='" & idxlist.List(klist.ListIndex) + "'"
currentk = idxlist.List(klist.ListIndex)
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  On Error Resume Next
  postanredek.text = trm(rtmp!postanrede)
  rrr = Err
  On Error GoTo 0
  For i% = sflds To nfldsk
    If Not IsNull(rtmp.Fields(i%)) Then kdat(i% - sflds).text = rtmp.Fields(i%)
  Next i%
  For i% = 0 To nfldska
    If Not IsNull(rtmp.Fields(11 + i%)) Then kadat(i%).text = rtmp.Fields(i% + 11)
  Next i%
  If trm(datf(0).text) <> "" Then Call form1.chkallnums(datf(0).text, idxlist.List(klist.ListIndex), "email", kdat(5).text)
  If Not form1.isfieldmissing("kontakt", "opttel") Then
    optktel.text = trm(rtmp!opttel)
  Else
    optktel.text = "database needs extension."
  End If
End If
Command5.Enabled = False
Label52.Caption = transe("inkl. Adresse")
BackColor = form1.cleancolor()
knt_sav.Enabled = False
adr_sav.Enabled = False
Anrede.text = form1.meineanrede(idxlist.List(klist.ListIndex))
Abrede.text = form1.meineabrede(idxlist.List(klist.ListIndex))
Command12.Enabled = True
Command36.Enabled = True
If klist.ListIndex >= 0 Then
  klistid = idxlist.List(klist.ListIndex)
Else
  klistid = "-1"
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT typ,wert FROM adresstyp where vid='" & datf(0).text & "' and kid='" & klistid + "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
List1.Clear
If Not rtmp.EOF Then
  While Not rtmp.EOF
    If Left(rtmp!typ, 4) <> "rel:" Then
      List1.AddItem transe(rtmp!typ) & ": " & rtmp!wert
    Else
      tabkz = form1.getusersetting("relabkz_" + Mid(trm(rtmp!typ), 5), Mid(trm(rtmp!typ), 5))
      List1b.AddItem tabkz + ":" + trm(rtmp!wert)
    End If
    rtmp.MoveNext
  Wend
End If
p1offs% = 0
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error Resume Next
tr = Dir(dn$ + "\ico\*.jpg")
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  While tr <> "" And tr <> klist.List(klist.ListIndex) + ".jpg"
    p1offs% = p1offs% + 1
    tr = Dir
  Wend
Else
  Call form1.dbg2f("Fehler bei tr(" + dn$ + ")")
End If
      kadat(2).ToolTipText = transe("Postleitzahl")
      kadat(3).ToolTipText = transe("Ort")
      If Not form1.isfieldmissing("opt_adresspool", "id") Then
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
        c$ = "select id,Beschreibung from opt_adresspool where vid='" + datf(0).text + "' and kid='" + klistid + "'"
        r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
        updsv = False
        If r.EOF Then
          updsv = True
        Else
          updsv = True
          While Not r.EOF And updsv
            If LCase(trm(r!Beschreibung)) = "standard" Then updsv = False
            r.MoveNext
          Wend
        End If
        If updsv Then Call svkadr("Standard")
      End If




c$ = "select * from dochist where adresse='" + datf(0).text + "' and kontakt='" + klistid + "' and doctyp='" + transe("Datenänderung") + "';"
Debug.Print c$
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
Label9.Caption = transe("geändert:")
If Not s.EOF Then
  Label9.Caption = Label9.Caption + " von " + trm(s!Owner)
  datf(8).text = datfromsql(word1(trm(s!erstellt))) + " " + word2bis(trm(s!erstellt))
End If

Call rlist4
Call mlist
Call kposbuttonset
BackColor = bg
knt_sav.Enabled = False
adr_sav.Enabled = False
If autokategorie$ <> "" And List1.Visible = True Then
  For i% = 0 To List1.ListCount - 1
    If InStr(List1.List(i%), autokategorie$) = 1 Then
      List1.ListIndex = i%
      Exit For
    End If
  Next i%
End If

End Sub

Sub dbg(l$)

'd2infile = "shwAdrDetail": d2insub = "dbg"
Call form1.dbg(l$)

End Sub


Public Sub addtyp(typ$)
Dim kid$, kontakt$, hid$, c$, i%

'd2infile = "shwAdrDetail": d2insub = "addtyp"
kid$ = datf(0).text
If kid$ <> "" Then
  If klist.ListIndex >= 0 Then
    kontakt$ = "'" + idxlist.List(klist.ListIndex) + "'"
  Else
    kontakt$ = "'-1'"
  End If
  If Not form1.sqlqry _
   ( _
    "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + kid$ + "','" + typ$ + "',NULL," + kontakt$ + ")" _
    ) Then Exit Sub

   If form1.getusersetting("adresstypwechsellog", "nein") = "ja" Then
     hid$ = form1.newid("dochist", "id", 19)
     c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
            hid$ & "','" & kid$ & "'," + kontakt$ + ",'Gruppenwechsel','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','neu in AdrGrp " + typ$ + "','Emaileingang')"
        Call form1.sqlqry(c$)
   End If
End If
For i% = 0 To List1.ListCount
  If typ$ = transe(List1.List(i%)) Then Exit Sub
Next i%
List1.AddItem transe(typ$)
Call rlist4
For i% = 0 To List1.ListCount
  If typ$ = transe(List1.List(i%)) Then
    List1.ListIndex = i%
    Exit For
  End If
Next i%

End Sub


Private Sub knt_delallow_Click()
If Check1.value = 0 Then
  Check1.value = 1
Else
  Check1.value = 0
End If

End Sub

Private Sub knt_dsl_Click()
Call Command12_Click
End Sub

Private Sub knt_dwn_Click()
Call Command44_Click
End Sub

Private Sub knt_sav_Click()
Call Command5_Click
End Sub

Private Sub knt_up_Click()
Call Command45_Click
End Sub

Private Sub knumsel_Click()
kdat(5).text = knumsel.text

End Sub

Private Sub knumsel_DropDown()
Dim c$, rtmp As ADODB.Recordset, rrr, id$, kid$

knumsel.Clear
id$ = datf(0).text
If id$ = "" Then Exit Sub
kid$ = kdat(0).text
If kid$ = "" Or kid$ = "-1" Then Exit Sub

c$ = "SELECT num FROM opt_allenummern where kid='" + kid$ + "' and numtyp='email' order by num"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
While Not rtmp.EOF
  If trm(kdat(5).text) <> trm(rtmp!num) Then knumsel.AddItem trm(rtmp!num)
  rtmp.MoveNext
Wend
knumsel.AddItem kdat(5).text
Call form1.chkallnums(id$, kid$, "email", kdat(5).text)

End Sub

Private Sub Label10_dblClick()
'd2infile = "shwAdrDetail": d2insub = "Label10_dblClick"
Call form1.dialme(kdat(3).text)
End Sub

Private Sub Label11_Click()
Dim brw$, X

'd2infile = "shwAdrDetail": d2insub = "Label11_Click"
Unload frmBrowser
DoEvents
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & mkhttp(datf(10).text), 1)
Else
  frmBrowser.StartingAddress = mkhttp(datf(10).text)
  Load frmBrowser
End If

End Sub


Private Sub Label16_dblClick()
'd2infile = "shwAdrDetail": d2insub = "Label16_dblClick"
Call form1.dialme(datf(9).text)
End Sub

Private Sub Label17_Click()
'd2infile = "shwAdrDetail": d2insub = "Label17_Click"
If gd1show.value = 0 Then
  gd1show.value = 1
Else
  gd1show.value = 0
End If

End Sub

Private Sub Label24_dblClick()
'd2infile = "shwAdrDetail": d2insub = "Label24_dblClick"
Call form1.dbg2f("shwadrdetail.label24_dblClick:kdat(6).Text=" + kdat(6).text)
Call form1.dialme(kdat(6).text)
End Sub

Private Sub Label27_Click()
Dim brw$, X

'd2infile = "shwAdrDetail": d2insub = "Label27_Click"
Unload frmBrowser
DoEvents
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & Chr$(34) & mkhttp(kdat(8).text) & Chr$(34), 1)
Else
  frmBrowser.StartingAddress = mkhttp(kdat(8).text)
  Load frmBrowser
End If



End Sub

Private Sub Label28_Click()
'd2infile = "shwAdrDetail": d2insub = "Label28_Click"
Call savecheck
Call Command2_Click

End Sub

Private Sub Label3_DblClick()

'd2infile = "shwAdrDetail": d2insub = "Label3_DblClick"

End Sub

Private Sub Label30_DblClick()

'd2infile = "shwAdrDetail": d2insub = "Label30_DblClick"

End Sub

Private Sub Label32_Click()
'd2infile = "shwAdrDetail": d2insub = "Label32_Click"
If gd1bez.value = 0 Then
  gd1bez.value = 1
Else
  gd1bez.value = 0
End If

End Sub

Private Sub Label4_DblClick()
Dim tz$, wert$

'd2infile = "shwAdrDetail": d2insub = "Label4_DblClick"
tz$ = form1.getusersetting("plzort-" & trm(datf(14)), "L P O")
wert$ = trm(InputBox("Wert von sysvar_system_plzort-" + trm(datf(14).text), "Land-PLZ-Ort Reihenfolge ändern", tz$))
If wert$ <> "" And wert$ <> tz$ Then
  tz$ = "delete from sysvars where owner='sysvar_system_plzort-" & trm(datf(14).text) & "'"
  Call form1.sqlqry(tz$)
  tz$ = "insert into sysvars (id,owner,wert) values ('" & _
    form1.newid("sysvars", "id", 8) & "','sysvar_system_plzort-" & trm(datf(14).text) & "','" & wert$ & "')"
  Call form1.sqlqry(tz$)
End If
End Sub

Private Sub Label41_DblClick()
'd2infile = "shwAdrDetail": d2insub = "Label41_DblClick"

End Sub

Private Sub Label45_DblClick()
'd2infile = "shwAdrDetail": d2insub = "Label45_DblClick"

End Sub

Private Sub Label46_Click()
'd2infile = "shwAdrDetail": d2insub = "Label46_Click"
If Check3.value = 1 Then
  Check3.value = 0
Else
  Check3.value = 1
End If

End Sub

Private Sub Label49_Click()
If intuse.value = 0 Then
  intuse.value = 1
Else
  intuse.value = 0
End If
End Sub

Private Sub Label5_dblClick()
Call form1.dialme(datf(4).text)
End Sub

Private Sub Label50_DblClick()
Dim ll$

ll$ = LCase(Label50.Caption)
If InStr(LCase(ll$), "tel") > 0 Then Call form1.dialme(opttel.text)
End Sub

Private Sub Label53_Click()
If stcky_igno Then Exit Sub
If stcky.value = 0 Then
  stcky.value = 1
Else
  stcky.value = 0
End If

End Sub

Private Sub Label6_Click()
Call savecheck
Call Command12_Click
Call Command2_Click

End Sub


Private Sub List1_Click()
If InStr(Command14.Caption, transe("Zusatz-Infos")) > 0 Then
  gd1show.value = 0
  Call gd1show_Click
  Call Command14_Click
Else
  Unload zusinf
  Unload bezlist
  Call rlist4
End If
End Sub

Private Sub List1_DblClick()
Dim typ$, t0$, wert$, vid$, kid$, kontakt$, s$, i%

'd2infile = "shwAdrDetail": d2insub = "List1_DblClick"
typ$ = transo(List1.List(List1.ListIndex))
t0$ = transo(List1.List(List1.ListIndex))
wert$ = ""
If InStr(typ$, ": ") Then
  wert$ = Mid$(typ$, InStr(typ$, ": ") + 2)
  typ$ = Left$(typ$, InStr(typ$, ":") - 1)
End If
vid$ = datf(0).text
If vid$ = "" Then Exit Sub
kid$ = datf(0).text
If kid$ <> "" Then
  If klist.ListIndex >= 0 Then
    kontakt$ = idxlist.List(klist.ListIndex)
  Else
    kontakt$ = "-1"
  End If
End If
wert$ = InputBox(transe("Wert von") + " " + typ$, transe("Wert ändern"), wert$)
List1.RemoveItem List1.ListIndex
s$ = ""
If wert$ <> "" Then
  s$ = "update adresstyp set wert='" + wert$ + "' where vid ='" + vid$ + "' and typ='" + typ$ + "' and kid='" + kontakt$ + "'"
  form1.sqlqry (s$)
  s$ = transe(typ$) + ": " + wert$
  List1.AddItem s$
Else
  s$ = transe(t0$)
  List1.AddItem s$
End If
If s$ <> "" Then
  For i% = 0 To List1.ListCount - 1
    If List1.List(i%) = s$ Then
      List1.ListIndex = i%
      Exit For
    End If
  Next i%
End If
End Sub

Sub rlist3()
Dim rtmp As ADODB.Recordset, o%, l$, i%
Dim inifile As String, c$, rrr

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "rlist3"
If fl_rl3% = 1 Then Exit Sub
fl_rl3% = 1
nl3fl% = 1
List3.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id FROM auftrittstypen order by sortierung"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While Not rtmp.EOF
  List3.AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend
List3.ListIndex = -1

inifile = form1.mylocaldatadir() + "\positions\" & Me.name & ".list3"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    For i% = 0 To List3.ListCount - 1
      If transo(List3.List(i%)) = l$ Then
        List3.Selected(i%) = True
        Exit For
      End If
    Next i%
  Wend
  Close #o%
End If

fl_rl3% = 0
nl3fl% = 0

End Sub

Private Sub List1b_Click()
'd2infile = "shwAdrDetail": d2insub = "List1b_Click"
If l1bdont Then Exit Sub
If InStr(Command14.Caption, transe("Zusatz-Infos")) > 0 Then
  gd1show.value = 0
  Call gd1show_Click
  Call Command14_Click
Else
  Call rlist4
End If

End Sub

Private Sub List1b_DblClick()
Dim i As Integer, sid$, sida$, sidk$, p As Integer

'd2infile = "shwAdrDetail": d2insub = "List1b_DblClick"
i = List1b.ListIndex
If i < 0 Then Exit Sub

p = InStr(List1b.List(i), ":")
If p < 1 Or p > Len(List1b.List(i)) - 1 Then Exit Sub

sid$ = trm(Mid(List1b.List(i), p + 1))
sida$ = sid$: sidk$ = ""
p = InStr(sida$, "{")
If p > 0 Then
  sidk$ = trm(Left(sid$, p - 1))
  sida$ = trm(Mid(sid$, p + 1)): sida$ = Left(sida$, Len(sida$) - 1)
End If
If Len(sida$) > 0 Then Call shwAdrDetail.refreshadrdetail(sida$, sidk$)

End Sub

Public Sub List2_DblClick()
Dim id$

'd2infile = "shwAdrDetail": d2insub = "List2_DblClick"
id$ = List2.List(List2.ListIndex)
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
Unload auftritt: Load auftritt
On Error Resume Next
Call auftritt.SetFocus
On Error GoTo 0
Call auftritt.showrec(id$, 0)

End Sub
Sub rlist4()
Dim r As ADODB.Recordset, s As ADODB.Recordset, i%, i2%, l1l$, l1bl$, cmd$, kid$
Dim lagerkennung As String, b1 As Double, b2 As Double, anz As Double, part$, vpos%, ccc$
Dim ksel$, rrr, sfddat$, sfwdat$, z, ad4$, n%, ad5$, l1i$, ad4a$, j%, j1%, zusi$
Dim kidx$

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "rlist4"
i% = List1.ListIndex
i2% = List1b.ListIndex
If i% < 0 Then
  List4.Clear
End If
'On Error GoTo errhdl
If Command14.Caption = transe("Zusatz-Infos") Then Exit Sub
List4.Clear
List5.Clear
List6.Clear
List10.Clear
l1l$ = ""
l1bl$ = ""
If Not List1b.Visible Then
  l1l$ = transo(List1.List(i%))
Else
  If i2% >= 0 Then l1bl$ = List1b.List(i2%)
End If
If InStr(l1l$, ":") > 0 Then l1l$ = Left(l1l$, InStr(l1l$, ":") - 1)
If klist.ListIndex >= 0 Then
  ksel$ = idxlist.List(klist.ListIndex)
  lagerkennung = form1.get_kontaktname_by_id(ksel$) & " {" & datf(0).text & "}"
Else
  ksel$ = ""
  lagerkennung = datf(0).text
End If
  If Not List1b.Visible Then

  zusi$ = form1.getusersetting("zusatzinfos", "normal")
  j1% = 0: vpos% = 240
  cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='" + l1l$ + "' ORDER BY typ, position"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    Call form1.dbg2f("zusatzinfos=erweitert", "shwadrdetail", "rlist4")
    If zusi$ = "erweitert" Then
      Unload zusinf
      Unload bezlist
      DoEvents
      On Error Resume Next
      Load zusinf
      Call zusinf.SetFocus
      On Error GoTo 0
    End If
  Else
    zusi$ = "normal"
    Unload zusinf
    Unload bezlist
  End If
  While Not r.EOF And break% = 0
    sfddat$ = "": sfwdat$ = ""
    z = r!zeilen
    ad4$ = transe(r!feldname) & ": "
    sfwdat$ = r!feldname
    cmd$ = "SELECT id,Felddaten From auftritthigru where auftrittstyp='" + l1l$ + "' and auftrittsid='" + datf(0).text + ksel$ + "' and feldname='" + r!feldname + "'"
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr <> 0 Then
      MousePointer = 0
      Exit Sub
    End If
    n% = -1
    If Not s.EOF Then
      ad5$ = "update auftritthigru set felddaten='$$$' where id='" + s!id + "'"
      If Not IsNull(s!felddaten) Then sfddat$ = s!felddaten
      n% = repl1310rc(sfddat$)
    Else
      l1i$ = transo(List1.List(i%))
      If InStr(l1i$, ":") > 0 Then l1i$ = Left$(l1i$, InStr(l1i$, ":") - 1)
      ad5$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + form1.newid("auftritthigru", "id", 18) + "','" + datf(0).text + ksel$ + "','" + l1i$ + "','" + r!feldname + "','$$$')"
      If InStr(form1.mustfield, l1i$ + "|" + trm(r!feldname)) > 0 Then
        ccc$ = strrepl(ad5$, "$$$", "")
Call form1.dbg2f(ccc$, "shwadrdetail", "rlist4")
        Call form1.sqlqry(ccc$)
        If ksel$ <> "" Then
          If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
            ccc$ = "update auftritthigru set opt_kid='" + ksel$ + "' where id='" + s!id + "'"
            Call form1.sqlqry(ccc$)
          End If
        End If
      End If
    End If
    ad4a$ = ad4$
    If zusi$ = "erweitert" And j1% < 15 And vpos% <= 5280 Then
        If vpos% = 240 Then
          zusinf.Caption = transo(cut_d1(List1.List(i%), ":")) + " " + datf(0).text
          zusinf.adrid.Caption = datf(0).text
          zusinf.kid.Caption = ksel$
          zusinf.typid.Caption = l1l$
          If ksel$ <> "" Then
            zusinf.Caption = zusinf.Caption + " Kontakt: " + kdat(2).text
          End If
        End If
        zusinf.Label1(j1%).Visible = True
        zusinf.Label1(j1%).Top = vpos%
        zusinf.Label1(j1%).Caption = ad4$
        zusinf.Text1(j1%).Top = vpos%
        zusinf.Text1(j1%).Visible = True
        zusinf.Text1(j1%).text = sfddat$
        zusinf.Text2(j1%).text = ad5$
        zusinf.Text1(j1%).Height = z * 285
        zusinf.Text1(j1%).BackColor = RGB(255, 255, 255)
        vpos% = vpos% + z * 285 + 75
        j1% = j1% + 1
    End If
    ad5$ = trm(z) + ":" + ad5$
    j% = 0
    While n% >= 0
      ad4a$ = ad4a$ + rcarr$(j%)
      If n% > 0 Then ad4a$ = ad4a$ + Chr$(13) + Chr$(10)
      j% = j% + 1
      n% = n% - 1
    Wend
    n% = repl1310rc(ad4a$)
    For j% = 0 To n%
      List5.AddItem ad5$
      List4.AddItem rcarr$(j%)
      List6.AddItem sfddat$
      List10.AddItem sfwdat$
    Next j%
    DoEvents
    r.MoveNext
  Wend
  r.Close
  If zusi$ = "erweitert" Then zusinf.Command4.Enabled = False
  Else    '1..visible
    If List1b.Visible Then
    
    List4.AddItem l1bl$
    Unload bezlist
    Unload zusinf
    DoEvents
    On Error Resume Next
    Load bezlist
    Call bezlist.SetFocus
    On Error GoTo 0
    kid$ = "-1": kidx$ = "-1"
    If klist.ListIndex >= 0 Then
      kid$ = klist.List(klist.ListIndex)
      kidx$ = idxlist.List(klist.ListIndex)
    End If
'    Call bezlist.setcurrent(idshow.Caption, kid$, kidx$)
    Call bezlist.currentrel(l1bl$)

    End If
  End If

Exit Sub
errhdl:
  rrr = Err
  If rrr <> 0 Then
    If rrr <> 3420 Then MsgBox transe("Fehler #") & rrr & " " + Error$(rrr)
    On Error GoTo 0
    break% = 1
    Unload adrselect
    Exit Sub
  End If
  Resume Next
End Sub

Private Sub List3_Click()
Dim inifile As String, o%, rrr, i%

'd2infile = "shwAdrDetail": d2insub = "List3_Click"
If nl3fl% = 1 Then Exit Sub
inifile = form1.mylocaldatadir() + "\positions\" & Me.name & ".list3"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  For i% = 0 To List3.ListCount - 1
    If List3.Selected(i%) Then Print #o%, transo(List3.List(i%))
  Next i%
  Close #o%
End If

End Sub

Private Sub List4_Click()
Dim i%, l1l$

'd2infile = "shwAdrDetail": d2insub = "List4_Click"
i% = List1.ListIndex
If i% < 0 Then Exit Sub
l1l$ = transo(List1.List(i%))
If InStr(l1l$, ":") > 0 Then l1l$ = Left(l1l$, InStr(l1l$, ":") - 1)
If LCase(l1l$) = "warenlager" Then Exit Sub

On Error Resume Next
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
On Error GoTo 0

End Sub
Private Function repl1310rc(l$)
Dim r$, n%, li$, i%, z$

'd2infile = "shwAdrDetail": d2insub = "repl1310rc"
n% = 0
rcarr$(n%) = ""
If InStr(l$, Chr$(13) + Chr$(10)) > 0 Then
  li$ = l$
  While InStr(li$, Chr$(13) + Chr$(10)) > 0
  For i% = 1 To Len(li$)
    z$ = Mid$(li$, i%, 1)
    If z$ = Chr$(13) Then
      i% = i% + 1
      n% = n% + 1
      rcarr$(n%) = ""
    Else
      rcarr$(n%) = rcarr$(n%) + z$
    End If
  Next i%
  li$ = r$
  Wend
Else
  rcarr$(n%) = l$
End If
repl1310rc = n%

End Function

Private Sub List4_DblClick()
Dim p1$, i%, l1l$, p As Integer, neuwert As String, neukwert As String, sid0 As String
Dim sid$, sida$, sidk$, neuawert, w%, params$

'd2infile = "shwAdrDetail": d2insub = "List4_DblClick"
If List1b.Visible Then
  i% = List1b.ListIndex
  If i% < 0 Then Exit Sub

  p = InStr(List1b.List(i%), ":")
  If p < 1 Or p > Len(List1b.List(i)) - 1 Then Exit Sub
  sid$ = trm(Mid(List1b.List(i%), p + 1))
  sid0$ = sid$
  sida$ = sid$: sidk$ = ""
  p = InStr(sida$, "{")
  If p > 0 Then
    sidk$ = trm(Left(sid$, p - 1))
    sida$ = trm(Mid(sid$, p + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If sidk$ <> "" Then sid$ = sidk$
  Load adrselect
  Call adrselect.sel_init(sid$, "")
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then
    neukwert = adrselect.get_kontsel()
    neuwert = adrselect.sel_getselected(): neuawert = neuwert
    If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
    If sid0$ <> neuwert Then Call addrel(List2.List(List2.ListIndex), neuwert)
    Unload adrselect
  End If
  Exit Sub
End If

w% = List4.ListIndex
If w% < 0 Then Exit Sub

i% = List1.ListIndex
If i% < 0 Then
  l1l$ = cut_d1(List4.List(w%), ":")
  Unload wweg: DoEvents
  On Error Resume Next
  Load wweg
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  On Error Resume Next
  Call wweg.SetFocus
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  wweg.Caption = "Warenbewegung: " & l1l$
  Call wweg.rereadorder
  Exit Sub
End If

l1l$ = transo(List1.List(i%))
If InStr(l1l$, ":") > 0 Then l1l$ = Left(l1l$, InStr(l1l$, ":") - 1)

p1$ = ""
params$ = List5.List(w%)
Load multilineinput
On Error Resume Next
Call multilineinput.SetFocus
On Error GoTo 0
If klist.ListIndex >= 0 Then
  multilineinput.Text3.text = "update kontakt set id='" + kdat(0).text + "' where id='" + kdat(0).text + "'"
  If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
    multilineinput.Text3.text = "update auftritthigru set opt_kid='" + kdat(0).text + "' where auftrittsid='" + idshow.Caption + kdat(0).text + "'"
  End If
Else
  multilineinput.Text3.text = "update adresse set id='" + idshow.Caption + "' where id='" + idshow.Caption + "'"
End If
Call multilineinput.init(params$)
Call multilineinput.setdeflt(List6.List(w%))
Call multilineinput.setcap(List10.List(w%))
DoEvents
If p1$ <> "" Then
  responder.hid.text = p1$
End If
End Sub
Public Sub reshow4()
'd2infile = "shwAdrDetail": d2insub = "reshow4"
Call rlist4
End Sub

Private Sub List7_Click()
'd2infile = "shwAdrDetail": d2insub = "List7_Click"
Label7.Caption = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text) + "\" + List7.List(List7.ListIndex)
End Sub

Private Sub List7_DblClick()
Dim fn$, X

'd2infile = "shwAdrDetail": d2insub = "List7_DblClick"
fn$ = Label7.Caption
Select Case LCase(Right$(fn$, 4))
  Case ".wav": Call sndPlaySound(fn$, SND_SYNC)
  Case ".mp3": 'x = Shell(form1.getmymp3player() + " " + Label7.Caption, 1)
  Case ".rtf": Call form1.openthisdoc(fn$, "")
  Case ".doc": Call form1.openthisdoc(fn$, "")
  Case ".txt": X = Shell("notepad.exe " & fn$, 1)
  Case Default:
End Select
End Sub

Private Sub List8_DblClick()
Dim i$

'd2infile = "shwAdrDetail": d2insub = "List8_DblClick"
i$ = List8.List(List8.ListIndex)
form1.sqlqry (i$)
List8.RemoveItem List8.ListIndex

End Sub

Private Sub List9_DblClick()
Dim i%, trgp$, vorlage$, eadr$

'd2infile = "shwAdrDetail": d2insub = "List9_DblClick"
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error GoTo 0

i% = List9.ListIndex
If i% >= 0 Then
  vorlage$ = List9.List(List9.ListIndex) + ".rtf"
  i% = klist.ListIndex
  MousePointer = 11
  trgp$ = ""
  If usempth.value = 1 Then
    On Error Resume Next
    MkDir form1.s0dir() + "\" + form1.medien() + "\"
    MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
    On Error GoTo 0
    trgp$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
  End If
  If i% < 0 Then
    eadr$ = "": If cadrpbez.text <> "Standard" Then eadr$ = "elseadr"
    Call form1.faxan(datf(0).text, "-1", vorlage$, "", "", trgp$, eadr$)
  Else
    Call form1.faxan(datf(0).text, idxlist.List(klist.ListIndex), vorlage$, "", "", trgp$, "")
  End If
  MousePointer = 0
End If
End Sub

Private Sub optktel_Change()
Command5.Enabled = True
knt_sav.Enabled = True
BackColor = form1.dirtycolor()
'telfaxhandy

kdat(7).text = onlynums(kdat(4).text) + " " + onlynums(kdat(3).text) + " " + onlynums(kdat(6).text) + " " + onlynums(optktel.text)

End Sub

Private Sub opttel_Change()
  datf(12).text = onlynums(datf(4).text) + " " + onlynums(datf(5).text) + " " + onlynums(datf(9).text) + " " + onlynums(opttel.text)
  Command4.Enabled = True
  adr_sav.Enabled = True
  BackColor = form1.dirtycolor()
End Sub

Private Sub opttel_DblClick()
If form1.darf_ich_sprechen() = True Then Call spknum(opttel.text())
End Sub

Private Sub opttel_GotFocus()
prv$ = opttel.text
End Sub

Private Sub p1_DblClick(Index As Integer)
'd2infile = "shwAdrDetail": d2insub = "p1_DblClick"
Call p1cmd_Click(Index)

End Sub

Private Sub p1cmd_Click(Index As Integer)
Dim dn$, tn$

'd2infile = "shwAdrDetail": d2insub = "p1cmd_Click"
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text) + "\"

tn$ = p1cmd(Index).Caption
If tn$ <> "" Then
  If exist(dn$ + tn$) = 0 Then
    tn$ = Left$(tn$, Len(tn$) - 4) + ".bmp"
  End If
  Load vwr
  vwr.SetFocus
  vwr.Caption = dn$ + tn$
  Call vwr.rdrw
End If
End Sub

Private Sub plzp_Change()
'd2infile = "shwAdrDetail": d2insub = "plzp_Change"
  Command4.Enabled = True
  adr_sav.Enabled = True
  BackColor = form1.dirtycolor()

End Sub

Private Sub plzp_KeyDown(KeyCode As Integer, Shift As Integer)
'd2infile = "shwAdrDetail": d2insub = "plzp_KeyDown"
Call esccntinc(KeyCode)
End Sub

Private Sub postanredea_Change()
'd2infile = "shwAdrDetail": d2insub = "postanredea_Change"

  Command4.Enabled = True
  adr_sav.Enabled = True
  BackColor = form1.dirtycolor()


End Sub

Private Sub postanredea_Click()
'd2infile = "shwAdrDetail": d2insub = "postanredea_Click"
Call postanredea_Change

End Sub

Private Sub postanredea_GotFocus()
If knt_sav.Enabled Then Call savecheck
End Sub

Private Sub postanredek_Change()
'd2infile = "shwAdrDetail": d2insub = "postanredek_Change"
Command5.Enabled = True
knt_sav.Enabled = True
BackColor = form1.dirtycolor()

End Sub

Private Sub postanredek_Click()
'd2infile = "shwAdrDetail": d2insub = "postanredek_Click"
Call postanredek_Change
End Sub

Private Sub postf_Change()
'd2infile = "shwAdrDetail": d2insub = "postf_Change"
  Command4.Enabled = True
  adr_sav.Enabled = True
  BackColor = form1.dirtycolor()

End Sub

Private Sub postf_KeyDown(KeyCode As Integer, Shift As Integer)
'd2infile = "shwAdrDetail": d2insub = "postf_KeyDown"
Call esccntinc(KeyCode)
End Sub

Private Sub prio_Change()
Dim c As String, id, p As String, nid As String

'd2infile = "tplan": d2insub = "prio_Change"
p = UCase(prio.text)
If p < "A" And p <> "" Then p = "A"
If p > "Z" And p <> "" Then p = "Z"
prio.text = p
id = trmx1(datf(0).text)
If id <> "" Then
  c = "delete from opt_prios where userid='" + form1.getuserid() + "' and evnt='A:" + id + "';"
  Call form1.sqlqry(c)
  If p <> "" Then
    nid = form1.newid("opt_prios", "id", 36)
    c = "insert into opt_prios (id,evnt,userid,prio) values('" + _
        nid + "','A:" + _
        id + "','" + _
        form1.getuserid() + "','" + _
         p + "');"
    Call form1.sqlqry(c)
  End If
  If form1.priosopen Then Call prios.Command20_Click
End If

End Sub

Private Sub repert_Click()
Dim id As String

Load repertoire
id = trmx1(datf(0).text)
If id <> "" Then

If kdat(2).text <> "" Then id = kdat(2).text + "{" + id + "}"
repertoire.artid.Caption = id
On Error Resume Next
Call repertoire.SetFocus
On Error GoTo 0

End If
End Sub

Private Sub season_Change()
Dim sd$, p%, FD$, td$

'd2infile = "shwAdrDetail": d2insub = "season_Change"
sd$ = season.text
p% = InStr(sd$, "-")
If p% > 1 Then
  On Error Resume Next
  FD$ = trm(Left(sd$, p% - 1))
  td$ = trm(Mid(sd$, p% + 1))
  Text1.text = Date - CDate(FD$)
  Text2.text = CDate(td$) - Date
  On Error GoTo 0
End If
End Sub

Private Sub season_Click()
'd2infile = "shwAdrDetail": d2insub = "season_Click"
Call season_Change
End Sub

Private Sub stcky_Click()
Dim c$

If stcky_igno Then Exit Sub
If trm(idshow.Caption) = "" Then
  stcky.value = 0
  Exit Sub
End If
If stcky.value = 0 Then
  c$ = "0"
  c$ = "delete from sysvars where owner='sysvar_" + form1.getuserid() + "_zzzadr_sticky_" + trm(idshow.Caption) + "'"
  Call form1.sqlqry(c$)
Else
  c$ = "1"
  Call form1.setusersetting("zzzadr_sticky_" + trm(idshow.Caption), "1")
End If

End Sub

Private Sub suchen_DblClick()
'd2infile = "shwAdrDetail": d2insub = "suchen_DblClick"

End Sub

Private Sub Text1_Change()
'd2infile = "shwAdrDetail": d2insub = "Text1_Change"
toffsvon = Val(Text1.text)
Call form1.setmylastFormVar(Me.name, "toffsvon", Text1.text)
End Sub

Private Sub Text1_DblClick()
'd2infile = "shwAdrDetail": d2insub = "Text1_DblClick"
  With frmCalendar
    .init Text1, Text1.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar
  Call form1.dbg2f("cdate: " + trm(Date) + ", text: " + trm(Text1.text))
  If Trim(nonums(trm(Text1.text))) = "" Or Trim(nonums(trm(Text1.text))) = "-" Then Exit Sub
  Text1.text = CDate(datum2sql("" & Date)) - CDate(datum2sql(Text1.text))
End Sub

Private Sub Text2_Change()
'd2infile = "shwAdrDetail": d2insub = "Text2_Change"
toffsbis = Val(Text2.text)
Call form1.setmylastFormVar(Me.name, "toffsbis", Text2.text)
End Sub

Sub mlist()
Dim tr, t$, i%, dn$, rrr, o%, l$

'd2infile = "shwAdrDetail": d2insub = "mlist"
If Height < 8000 Then Exit Sub

p1cmdmax% = 3

MousePointer = 11
For i% = 0 To p1cmdmax%
  p1(i%).Picture = pnull.Picture
  p1cmd(i%).Caption = ""
Next i%
'inhalt jpg und bmp
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
i% = 0
tr = Dir(dn$ + "\*.jpg")
While tr <> "" And i% <= p1cmdmax% + p1offs%
  If i% >= p1offs% And (i% - p1offs%) <= p1cmdmax% Then
    rrr = 0
    On Error Resume Next
    p1(i% - p1offs%).Picture = LoadPicture(dn$ + "\" + tr)
    DoEvents
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then p1(i% - p1offs%).Picture = pnull.Picture
    p1cmd(i% - p1offs%).Caption = tr
  End If
  i% = i% + 1
  On Error Resume Next
  tr = Dir
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then tr = ""
Wend
tr = Dir(dn$ + "\*.bmp")
While tr <> "" And i% <= p1cmdmax% + p1offs%
  If i% >= p1offs% And (i% - p1offs%) <= p1cmdmax% Then
    rrr = 0
    On Error Resume Next
    p1(i% - p1offs%).Picture = LoadPicture(dn$ + "\" + tr)
    DoEvents
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then p1(i% - p1offs%).Picture = pnull.Picture
    p1cmd(i% - p1offs%).Caption = tr
  End If
  i% = i% + 1
  tr = Dir
Wend
'inhalt rest
List7.Clear
tr = Dir(dn$ + "\*.*")
While tr <> ""
  List7.AddItem tr
  tr = Dir
Wend
List8.Clear
If exist(dn$ + "\sql.cmd") Then
  o% = FreeFile
  Open dn$ + "\sql.cmd" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    List8.AddItem l$
    DoEvents
  Wend
  Close #o%
  Kill dn$ + "\sql.cmd"
  While List8.ListCount > 0
    form1.sqlqry (List8.List(0))
    List8.RemoveItem 0
    DoEvents
  Wend
End If

If Width >= 9000 Then
  List9.Clear
  tr = Dir(form1.vorlagenverzeichnis() + "\adresse_*.rtf")
  rrr = Err
  On Error GoTo 0
  While tr <> "" And rrr = 0
    t$ = tr
    List9.AddItem basename(t$, ".rtf")
    tr = Dir
  Wend
End If
MousePointer = 0
End Sub
Public Sub savecheck()
Dim antw As Integer
'd2infile = "shwAdrDetail": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    'Call Command4_Click
    Call Command5_Click
    Command36.Enabled = False
  End If
End If
BackColor = form1.cleancolor()
knt_sav.Enabled = False
adr_sav.Enabled = False
End Sub

Private Sub Text2_DblClick()
'd2infile = "shwAdrDetail": d2insub = "Text2_DblClick"

  With frmCalendar
    .init Text2, Text2.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text2.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar
  Call form1.dbg2f("cdate: " + trm(Date) + ", text: " + trm(Text2.text))
  If Trim(nonums(trm(Text2.text))) = "" Or Trim(nonums(trm(Text2.text))) = "-" Then Exit Sub
  Text2.text = CDate(datum2sql(Text2.text)) - CDate(datum2sql("" & Date))

End Sub

Private Sub usempth_Click()
'd2infile = "shwAdrDetail": d2insub = "usempth_Click"
Call form1.setmylastFormVar(Me.name, "usempth", trm(usempth.value))

On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(datf(0).text)
On Error GoTo 0

End Sub
Sub rcombo2()
Dim rtmp As ADODB.Recordset
Dim seli%, cmd$, dv$, db$, rrr

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "rcombo2"
cmd$ = "SELECT id,von,bis FROM tplan where "

'vergangene nicht, oder?
dv$ = datum2sql(Date - toffsvon)
'dv$ = datum2sql(Date - 0)
db$ = datum2sql(Date + toffsbis)
cmd$ = cmd$ + "((von>='" + dv$ + "') and (von<='" + db$ + "')) "
cmd$ = cmd$ + "ORDER BY von"

Combo2.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  Combo2.AddItem rtmp!id
  'Combo2.AddItem rtmp!id & " (" & rtmp!von & "-" & rtmp!bis & ")"
  rtmp.MoveNext
Wend
Combo2.text = ""

End Sub
Private Sub icalex_Click()
Dim ical$, icalo%, kid$

'd2infile = "shwAdrDetail": d2insub = "icalex_Click"
rlist2icalmode = True
kid$ = trm(datf(0).text)
If klist.ListIndex >= 0 Then
  kname$ = kdat(2).text
  kid$ = kdat(2).text + " {" + kid$ + "}"
End If
ical$ = form1.s0dir() & "\" & form1.medien() & "\" & kid$ & ".ics"
icalo% = FreeFile
Open ical$ For Output As #icalo%
Print #icalo%, "BEGIN:VCALENDAR"
Print #icalo%, "VERSION:2.0"
Print #icalo%, "PRODID:-//Agencyprof.de//NONSGML Agencyprof Calendar V0.3//EN"
Print #icalo%, "METHOD:PUBLISH"
Close #icalo%
Call rlist2
icalo% = FreeFile
Open ical$ For Append As #icalo%
Print #icalo%, "END:VCALENDAR"
Close #icalo%
Label36.Caption = form1.inmylanguage("Termine: ") + trm(List2.ListCount)
DoEvents
rlist2icalmode = False

End Sub

Sub esckcntinc(kc As Integer)
'd2infile = "shwAdrDetail": d2insub = "esckcntinc"
If kc = 27 Then
  esckcnt = esckcnt + 1
  If esckcnt > 1 Then
    esckcnt = 0
    Call Label41_DblClick
  End If
Else
  esckcnt = 0
End If

End Sub

Function plzortcheck(land$, plz$, ort$) As String
Dim p$, o$

'd2infile = "shwAdrDetail": d2insub = "plzortcheck"
plzortcheck = ""
If land$ = "" Or land$ = "D" Or land$ = transe("Deutchland") Then
  If plz$ = "" Then
    p$ = word1(ort$)
    If isnumber(p$) Then
      o$ = word2bis(ort$)
      plzortcheck = trm(p$) + "|" + trm(o$)
      Exit Function
    End If
  End If
End If

End Function

Sub importcsv(f As String)
Dim o As Integer, l As String, fld As String, mi As Integer, md As Integer, n As Integer
Dim fldn As String, currid As String, cmd As String, usekontakt As Boolean, i As Integer
Dim anr As String, nam As String, ask As Integer

'd2infile = "shwAdrDetail": d2insub = "importcsv"
mi = impfelder.Top
md = impdaten.Top
impfelder.Clear
impfelder.Top = Shape3.Top
impfelder.Left = Shape3.Left
DoEvents
usekontakt = False
o = FreeFile
Open f For Input As #o
Line Input #o, l
Do
  fld = LCase(cut_d1(l, ";"))
  If fld = "telefon" Then fld = "tel"
  If fld = "kontakt" Then usekontakt = True
  l = cut_d2bis(l, ";")
  impfelder.AddItem fld
Loop Until l = ""
ask = MsgBox("Diese Felder werden importiert:", vbYesNo + vbCritical + vbDefaultButton2, "Felder importieren?")
If ask = vbYes Then
  MousePointer = 11: DoEvents
  impfelder.Top = mi
  impdaten.Top = md
  impdaten.Clear
  impdaten.Top = Command5.Top
  impdaten.Left = Command5.Left
  DoEvents
  While Not EOF(o)
    Line Input #o, l
    impdaten.AddItem l
  Wend
  While impdaten.ListCount > 0
    impdaten.ListIndex = 0
    l = impdaten.List(0)
    n = 0
    DoEvents
    Do
      fld = cut_d1(l, ";")
      l = cut_d2bis(l, ";")
      fldn = impfelder.List(n)
      Select Case fldn
        Case "gruppe":
                    fld = trmvalidate(strrepl(fld, "/", ","))
                    fld = strrepl(fld, "+", ",")
                    While fld <> ""
                      cmd = trm(cut_d1(fld, ","))
                      fld = trm(cut_d2bis(fld, ","))
                      For i = 0 To List1.ListCount - 1
                        If InStr(List1.List(i), cmd + ":") = 1 Then Exit For
                      Next i
                      If i >= List1.ListCount Then
                        form1.sqlqry ("insert into adresstypen (id) values('" & cmd & "')")
                        Call addtyp(cmd)
                      End If
                    Wend
        Case "kontakt":
                   klist.ListIndex = -1
                   For i = 0 To klist.ListCount - 1
                     If InStr(klist.List(i), fld) > 0 Or InStr(fld, klist.List(i)) > 0 Then
                       klist.ListIndex = i
                       DoEvents
                       Exit For
                     End If
                   Next i
                   If klist.ListIndex < 0 Then
                     Call Command7_Click
                     DoEvents
                     anr = ""
                     If Left(fld, 4) = "Herr" Or Left(fld, 4) = "Frau" Then
                       anr = word1(fld)
                       nam = word2bis(fld)
                     End If
                     postanredek.text = anr
                     kdat(2).text = nam
                     Call Command5_Click
                     DoEvents
                     klist.ListIndex = -1
                     For i = 0 To klist.ListCount - 1
                       If InStr(klist.List(i), fld) > 0 Or InStr(fld, klist.List(i)) > 0 Then
                         klist.ListIndex = i
                         DoEvents
                         Exit For
                       End If
                     Next i
                   End If
        Case "id": currid = fld
                   cmd = "insert into adresse (id,name) values('" + currid + "','" + currid + "');"
                   Call form1.sqlqry(cmd)
                   Call refreshadrdetail(currid, "")
                   DoEvents
        Case Else:
                    If usekontakt Then
                      cmd = "update kontakt set " + fldn + "='" + fld + "' where id='" + idxlist.List(klist.ListIndex) + "';"
                      Call form1.sqlqry(cmd)
                    Else
                      cmd = "update adresse set " + fldn + "='" + fld + "' where id='" + currid + "';"
                      Call form1.sqlqry(cmd)
                    End If
      End Select
      n = n + 1
    Loop Until l = ""
    impdaten.RemoveItem 0
  Wend
  MousePointer = 0: DoEvents
End If
Close #o
impfelder.Top = mi
impdaten.Top = md

End Sub

Public Sub addrel(typ$, adr As String)
Dim kid$, kontakt$, hid$, c$, invtyp As String, cmd$, rrr, tabkz As String
Dim sida$, sidk$, sidp%, rtmp As ADODB.Recordset, invwert As String, sid$, i%

Dim d2infile As String, d2insub As String
d2infile = "shwAdrDetail": d2insub = "addrel"

tabkz = form1.getusersetting("relabkz_" + typ$, typ$)
List1b.AddItem typ$ + ":" + adr
kid$ = datf(0).text
If kid$ <> "" Then
  If klist.ListIndex >= 0 Then
    kontakt$ = "'" + idxlist.List(klist.ListIndex) + "'"
  Else
    kontakt$ = "'-1'"
  End If
  form1.sqlqry _
   ( _
    "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + kid$ + "','rel:" + typ$ + "','" + adr + "'," + kontakt$ + ")" _
   )
  If form1.getusersetting("adresstypwechsellog", "nein") = "ja" Then
    hid$ = form1.newid("dochist", "id", 19)
    c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
            hid$ & "','" & kid$ & "'," + kontakt$ + ",'Gruppenwechsel','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','neu in AdrGrp " + typ$ + "','Emaileingang')"
    Call form1.sqlqry(c$)
  End If
  invtyp = form1.getusersetting("inversrelation_" + typ$, "")
  If invtyp <> "" Then
    sidk$ = "-1"
    sida$ = adr
    invwert = kid$
    If kontakt$ <> "-1" And kontakt$ <> "'-1'" Then
      invwert = form1.get_kontaktname_by_id(kontakt$) + " {" + invwert + "}"
    End If
    If InStr(adr, "{") > 0 Then
      sid$ = adr
      sidp% = InStr(sid$, "{")
      sida$ = sid$
      sidk$ = trm(Left(sid$, sidp% - 1))
      sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
      sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
    End If
    cmd$ = "select * from adresstyp where vid='" + sida$ + "' and kid='" + sidk$ + "' and typ='rel:" + invtyp + "' and wert='" + invwert + "'"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rtmp.EOF Then
      form1.sqlqry _
       ( _
        "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + _
         sida$ + "','rel:" + invtyp + "','" + invwert + "','" + sidk$ + "')" _
       )
    End If
  End If
End If
If Not List1b.Visible Then Call Command42_Click
For i% = 0 To List1b.ListCount
  If typ$ = transe(List1b.List(i%)) Then
    List1b.ListIndex = i%
    Exit For
  End If
Next i%

End Sub

Sub kposbuttonset()

If Not usekpos Or klist.ListIndex < 0 Then
  Command44.Enabled = False
  Command45.Enabled = False
  knt_up.Enabled = False
  knt_dwn.Enabled = False
Else
  Command44.Enabled = True
  Command45.Enabled = True
  knt_up.Enabled = True
  knt_dwn.Enabled = True
End If

End Sub

Function gettyp(typ$) As Integer
Dim i%

gettyp = -1
For i% = 0 To 99
  If katid%(i%) < 0 Then Exit Function
  If katnames$(i%) = typ$ Then
    gettyp = katid%(i%)
    Exit Function
  End If
Next i%
End Function
