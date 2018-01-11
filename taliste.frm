VERSION 5.00
Begin VB.Form taliste 
   BackColor       =   &H00E0E0E0&
   Caption         =   "   Tournee-Angebote - AgencyProf"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   ForeColor       =   &H00000000&
   Icon            =   "taliste.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows-Standard
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
      Left            =   9960
      TabIndex        =   78
      ToolTipText     =   "Link entfernen"
      Top             =   480
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
      Left            =   9600
      TabIndex        =   79
      ToolTipText     =   "Link hinzufügen"
      Top             =   480
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   5520
      Left            =   10560
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   76
      ToolTipText     =   "Arten von Angebotslisten"
      Top             =   600
      Width           =   1695
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
      Left            =   600
      TabIndex        =   75
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   4800
      Picture         =   "taliste.frx":01CA
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Zum Projekt"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   4800
      Picture         =   "taliste.frx":0237
      Style           =   1  'Grafisch
      TabIndex        =   72
      ToolTipText     =   "Zum Projekt"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   4800
      Picture         =   "taliste.frx":02A4
      Style           =   1  'Grafisch
      TabIndex        =   71
      ToolTipText     =   "Zum Projekt"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   3000
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   25
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   69
      Text            =   "taliste.frx":0311
      ToolTipText     =   "Interne Anmerkung"
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   24
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   68
      Text            =   "taliste.frx":0317
      ToolTipText     =   "Anmerkung, die mitgedruckt wird"
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   7560
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   6120
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   21
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   65
      Text            =   "taliste.frx":031D
      ToolTipText     =   "scrollen, wenn nicht alle sichtbar"
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   3600
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   3600
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Index           =   18
      Left            =   3120
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "a&lle zeigen"
      Height          =   495
      Left            =   13200
      Style           =   1  'Grafisch
      TabIndex        =   52
      ToolTipText     =   "Markierung abwählen"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   13920
      Picture         =   "taliste.frx":0323
      Style           =   1  'Grafisch
      TabIndex        =   51
      ToolTipText     =   "löschen"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   12600
      Picture         =   "taliste.frx":0813
      Style           =   1  'Grafisch
      TabIndex        =   50
      ToolTipText     =   "Neue Angebotsliste"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Angebotslisten      >>"
      Height          =   255
      Left            =   8520
      TabIndex        =   48
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "taliste.frx":0BA5
      Style           =   1  'Grafisch
      TabIndex        =   47
      ToolTipText     =   "Schließen"
      Top             =   6000
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   12600
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   46
      ToolTipText     =   "Arten von Angebotslisten"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2160
      Picture         =   "taliste.frx":0DF5
      Style           =   1  'Grafisch
      TabIndex        =   45
      ToolTipText     =   "Wiedervorlage"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   960
      Picture         =   "taliste.frx":1174
      Style           =   1  'Grafisch
      TabIndex        =   44
      ToolTipText     =   "Speichern"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
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
      Left            =   8400
      Picture         =   "taliste.frx":151B
      Style           =   1  'Grafisch
      TabIndex        =   42
      ToolTipText     =   "Alarm setzen"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1560
      Picture         =   "taliste.frx":17F8
      Style           =   1  'Grafisch
      TabIndex        =   41
      ToolTipText     =   "Drucken"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5520
      TabIndex        =   40
      ToolTipText     =   "Adressdetails"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5520
      TabIndex        =   39
      ToolTipText     =   "Adressdetails"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   3000
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   16
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "taliste.frx":1C98
      ToolTipText     =   "Interne Anmerkung"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   15
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "taliste.frx":1C9E
      ToolTipText     =   "Anmerkung, die mitgedruckt wird"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   7440
      TabIndex        =   31
      Text            =   "Text1"
      ToolTipText     =   "Enddatum"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   6000
      TabIndex        =   29
      Text            =   "Text1"
      ToolTipText     =   "Anfangsdatum"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   12
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "taliste.frx":1CA4
      ToolTipText     =   "scrollen, wenn nicht alle sichtbar"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Index           =   9
      Left            =   3120
      TabIndex        =   22
      Text            =   "Text1"
      ToolTipText     =   "ID der Tournee"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   7
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "taliste.frx":1CAA
      ToolTipText     =   "Interne Anmerkung"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   6
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "taliste.frx":1CB0
      ToolTipText     =   "Anmerkung, die mitgedruckt wird"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   14
      Text            =   "Text1"
      ToolTipText     =   "Enddatum"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   6000
      TabIndex        =   12
      Text            =   "Text1"
      ToolTipText     =   "Anfangsdatum"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   3
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Text            =   "taliste.frx":1CB6
      ToolTipText     =   "scrollen, wenn nicht alle sichtbar"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "ID der Tournee"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      Picture         =   "taliste.frx":1CBC
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Neue Tournee"
      Top             =   240
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Liste der Tourneen"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox chgs 
      Height          =   3765
      Left            =   2880
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Listentyp"
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
      Left            =   10680
      TabIndex        =   77
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   10440
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tourneen"
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
      Left            =   960
      TabIndex        =   74
      ToolTipText     =   "Liste der Tourneen"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   26
      Left            =   2400
      TabIndex        =   61
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   25
      Left            =   6000
      TabIndex        =   60
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   24
      Left            =   5880
      TabIndex        =   59
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   23
      Left            =   7080
      TabIndex        =   58
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   22
      Left            =   5760
      TabIndex        =   57
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   21
      Left            =   2760
      TabIndex        =   56
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   20
      Left            =   2880
      TabIndex        =   55
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   19
      Left            =   2880
      TabIndex        =   54
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   18
      Left            =   2880
      TabIndex        =   53
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Angebotslisten"
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
      Left            =   12840
      TabIndex        =   49
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   38
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   16
      Left            =   6000
      TabIndex        =   36
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   15
      Left            =   5880
      TabIndex        =   34
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   6960
      TabIndex        =   32
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   30
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   28
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   11
      Left            =   2880
      TabIndex        =   26
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   10
      Left            =   2880
      TabIndex        =   24
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   9
      Left            =   2880
      TabIndex        =   21
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   6960
      TabIndex        =   13
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF2900&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Doppelklick zum Auswählen"
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   12480
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5775
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Innen ausgefüllt
      Height          =   2055
      Left            =   2760
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   7575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   2760
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2160
      Width           =   7575
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Innen ausgefüllt
      Height          =   2055
      Left            =   2760
      Shape           =   4  'Gerundetes Rechteck
      Top             =   4200
      Width           =   7575
   End
End
Attribute VB_Name = "taliste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prvw$, preselnodo%, csi%, nochg%

Private Sub Command1_Click()

'd2infile = "taliste": d2insub = "Command1_Click"
Hide
Unload taliste

End Sub

Private Sub Command10_Click(Index As Integer)

Dim idx%
'd2infile = "taliste": d2insub = "Command10_Click"
idx% = List1.ListIndex
If idx% < 0 Then Exit Sub
idx% = idx% + Index
If idx% > List1.ListCount - 1 Then Exit Sub
List1.ListIndex = idx%
DoEvents
Call List1_DblClick

End Sub

Private Sub Command11_Click()

'd2infile = "taliste": d2insub = "Command11_Click"
Call Command1_Click

End Sub

Private Sub Command12_Click()

'd2infile = "taliste": d2insub = "Command12_Click"
If Right(Command12.Caption, 2) = ">>" Then
  Width = 13365
  Command12.Caption = "<< " + transe("Angebotslisten")
Else
  Width = 11340
  Command12.Caption = transe("Angebotslisten") + " >>"
End If

End Sub

Private Sub Command13_Click()
Dim neuid As String

'd2infile = "taliste": d2insub = "Command13_Click"
neuid = InputBox(transe("Neuer Angebotstyp"), "Neuer Typ")

Call form1.sqlqry("INSERT INTO talisttyp (ID) VALUES('" & neuid & "')")
Call rlist2


End Sub

Public Sub Command15_Click()
Dim i%

'd2infile = "taliste": d2insub = "Command15_Click"
List1.ListIndex = -1
DoEvents
For i% = 0 To List2.ListCount - 1
  preselnodo% = 1: List2.Selected(i%) = False
Next i%
preselnodo% = 0
Call rlist1

End Sub

Private Sub Command18_Click()

'd2infile = "taliste": d2insub = "Command18_Click"
Call form1.handbuchcall("11-Tourneeangebote.htm")

End Sub

Private Sub Command2_Click()
Dim neuid As String, i%, nid$

'd2infile = "taliste": d2insub = "Command2_Click"
neuid = InputBox(transe("Neue Tournee"), transe("Neue Tournee"))

Call form1.sqlqry("INSERT INTO taliste (ID,von) VALUES('" & neuid & "','" & datum2sql(Date) & "')")
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = True Then
    nid$ = form1.newid("talisted", "id", 18)
    Call form1.sqlqry("insert into talisted (id,talistid,taid) values('" + nid$ + "','" + List2.List(i%) + "','" + neuid + "')")
  End If
Next i%
'besser später, wenn daten da
'Call Form1.sqlqry("INSERT INTO tplan (ID,von) VALUES('" & neuid & "','" & Date & "')")
Call rlist1

End Sub

Private Sub Command21_Click()
Dim tpid$, wert$, p%, cmd$, n$

tpid$ = Text1(0).text: If trm(tpid$) = "" Then Exit Sub
wert$ = "Bezeichnung=Linkziel"
wert$ = InputBox(transe("Neuer Link:"), transe("Link festlegen"), wert$)
p% = InStr(wert$, "=")
If p% > 0 Then
  n$ = trm(Mid$(wert$, p% + 1))
  wert$ = trm(Left$(wert$, p% - 1))
  cmd$ = "delete from auftritthigru where id='" + tpid$ + " " + wert$ + "' and auftrittstyp='taliste'"
  Call form1.sqlqry(cmd$)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values('"
  cmd$ = cmd$ + tpid$ + " " + wert$ + "','" + tpid$ + "','taliste','link_" + wert$ + "','" + n$ + "')"
  Call form1.sqlqry(cmd$)
  Call rgd2
Else
  MsgBox ("Syntaxfehler: Linkbezeichnung=Linkziel")
End If

End Sub

Private Sub Command22_Click()
  Set lvitem = gd2.SelectedItem
  id$ = lvitem.SubItems(2)
  c$ = "delete from auftritthigru where id='" & id$ & "' and auftrittstyp='taliste'"
  Call form1.sqlqry(c$)
  Call rgd2

End Sub

Private Sub Command3_Click()

'd2infile = "taliste": d2insub = "Command3_Click"
If trm(Text1(2).text) = "" Then
  Call Label1_DblClick(2)
Else
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(form1.getidbyname(Text1(2).text), "")
  Call shwAdrDetail.SetFocus
End If

End Sub

Private Sub Command4_Click()

'd2infile = "taliste": d2insub = "Command4_Click"
If trm(Text1(1).text) = "" Then
  Call Label1_DblClick(1)
Else
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(form1.getidbyname("" & Text1(1).text & ""), "")
  Call shwAdrDetail.SetFocus
End If

End Sub


Private Sub Command6_Click()
Dim o%, p%, nam$, vorlage$, extsel$, cmd$, fqvd$, fn$, fe$, ps%, php$, ph%, l$
Dim rtmp As ADODB.Recordset, talistfont As String, talistfontsize As String, tal As String
Dim stmp As ADODB.Recordset, sais$
Dim ttmp As ADODB.Recordset
Dim bkmstart$, bkmend$, ss%, sd%, tm%, ty%, smcount%, scount%, psais As String, csais As String
Dim q%, la$, l1$, iv$, t$, rev$, ttest$, ftm$, ln$, pb%, psaisb As String, csaisb As String

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "Command6_Click"
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
talistfont = "\f62"
talistfontsize = "\fs24"
tal = form1.getusersetting("talistfont")
If tal <> "" Then talistfont = tal
tal = form1.getusersetting("talistfontsize")
If tal <> "" Then talistfontsize = tal
vorlage$ = basename(form1.meinetalistevorlage(), ".rtf")
p% = 0
While p% < List2.ListCount
  If List2.Selected(p%) Then
    extsel$ = List2.List(p%)
    vorlage$ = vorlage$ + "-" & extsel$
    p% = List2.ListCount
  End If
  p% = p% + 1
Wend
fqvd$ = form1.vorlagenverzeichnis() + "\" & vorlage$ & ".rtf"
fn$ = form1.myuniquedocname("")
If fn$ = "" Then
  taliste.MousePointer = 0
  Exit Sub
End If
fe$ = DirName(fn$) & "\" & basename(fn$, ".rtf") & "_engl.rtf"
php$ = vorlage$: If InStr(php$, "-") > 0 Then php$ = Mid$(php$, InStr(php$, "-") + 1)
php$ = form1.s0dir() & "\" & php$

If exist(fqvd$) = 0 Then
  MsgBox "Vorlage unbekannt: " + fqvd$
  GoTo engver
End If
taliste.MousePointer = 11
DoEvents
ss% = Val("0" & trm(form1.getusersetting("saisonstart")))
sd% = Val("0" & trm(form1.getusersetting("saisondauer")))
If sd% = 0 Then sd% = 1
On Error Resume Next
Kill fn$ + ".pbk"
On Error GoTo 0
o% = FreeFile
Open fqvd$ For Input As #o%
p% = FreeFile
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  q% = InStr(l$, "TALISTE ")
  If q% > 0 Then
    ph% = FreeFile
    Open form1.s0dir() + "\oliste.txt" For Output As #ph%
    'cmd$ = "SELECT taliste.*, opt_talisted1.talistid as talistid1, talisted.talistid as listtyp FROM (taliste INNER JOIN talisted ON taliste.id = talisted.taid) INNER JOIN opt_talisted1 ON taliste.id = opt_talisted1.taid"
    'cmd$="SELECT taliste.*, opt_talisted1.talistid as talistid1, talisted.talistid as listtyp FROM (taliste INNER JOIN talisted ON taliste.id = talisted.taid) INNER JOIN opt_talisted1 ON taliste.id = opt_talisted1.taid"
    cmd$ = "SELECT taliste.*, talisted.talistid as listtyp FROM (taliste INNER JOIN talisted ON taliste.id = talisted.taid)"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
    Print #ph%, "orchester;solisten;dirigent;von;bis;anmerkung;typ;listtyp;saison;id;link1;link2;link3"
    While Not rtmp.EOF
      psais = ""
      cmd$ = "select talistid from opt_talisted1 where taid='" + trm(rtmp!id) + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      stmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
      While Not stmp.EOF
        psais = psais + trm(stmp!talistid) + ", "
        stmp.MoveNext
      Wend
      psais = trm(psais)
      While Right(psais, 1) = ",": psais = Left(psais, Len(psais) - 1): Wend
      If psais = "" Then psais = " "
      sais$ = form1.saison(datfromsql(rtmp!von))
      Print #ph%, "'" & rtmp!orchester & "';'" + strrepl(trm(rtmp!solisten), vbCrLf, "<br>");
      Print #ph%, "';'" & rtmp!dirigent & "';'" & datfromsqlshort(rtmp!von) & "';'" & datfromsqlshort(rtmp!bis) & "';'" & strrepl(trm(rtmp!anmerkung), vbCrLf, "<br>") & "';'" & psais & "';'" & trm(rtmp!listtyp) & "';'" & sais$ & "';'" & trm(rtmp!id) & "'";
      cmd$ = "select * from auftritthigru where auftrittsid='" + trm(rtmp!id) + "' and auftrittstyp='taliste' and instr(feldname,'link_')=1 order by id"
      Set ttmp = New ADODB.Recordset
      ttmp.CursorLocation = adUseServer
      ttmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
      sais$ = ""
      While Not ttmp.EOF
        sais$ = sais$ + "'" + trm(ttmp!felddaten) + "';"
        ttmp.MoveNext
      Wend
      Print #ph%, ";"; sais$
      rtmp.MoveNext
    Wend
    Close #ph%
    psais = "": psaisb = ""
    cmd$ = "SELECT taliste.*, talisted.talistid FROM taliste INNER JOIN talisted ON taliste.id = talisted.taid WHERE (((talisted.talistid)=""" + extsel$ + """)) order by von;"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
    While Not rtmp.EOF
    '\cellx3000\cellx6070\cellx9140
'    If trm(rtmp!von) <> "" Then
    If trm(rtmp!bis) = "" Then
        csaisb = "Projekte auf Anfrage"
        pbk% = FreeFile
        Open fn$ + ".pbk" For Append As #pbk%
        If psaisb <> csaisb Then
          psaisb = csaisb
          Print #pbk%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
          Print #pbk%, "{\b" & talistfont & talistfontsize & " \par \cf13 "; form1.repl1310rtf(trm(psaisb)); "}";
          Print #pbk%, " \par \cell"
          Print #pbk%, form1.repl1310rtf(" "); "\cell"
          Print #pbk%, "{\b" & talistfont & talistfontsize; "} \par "; " "; "\cell"
          Print #pbk%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
        Print #pbk%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
        Print #pbk%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!orchester); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.instrumentvon(l1$)
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #pbk%, "\par "; l1$ & iv$;
          If Len(la$) = 0 Then
            la$ = l1$ & iv$
          Else
            la$ = la$ & ", " & l1$ & iv$
          End If
        Wend
        Print #pbk%, "\cell"
        Print #pbk%, form1.repl1310rtf("" & rtmp!dirigent & ""); "\cell"
        Print #pbk%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "}\cell"
        Print #pbk%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        Close #pbk%
    Else
        ty% = Val(Left(trm(rtmp!von), 4))
        tm% = Val(Mid(trm(rtmp!von), 6, 2))
        smcount% = tm% - ss%
        If smcount% < 0 Then
          ty% = ty% - 1
          smcount% = smcount% + sd%
        End If
        scount% = smcount% / sd%
        csais = "Saison " + trm(ty%)
        If ss% <> 1 Then csais = csais + " / " + trm(ty% + 1)
        If psais <> csais Then
          psais = csais
          Print #p%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
          Print #p%, "{\b" & talistfont & talistfontsize & " \par \cf13 "; form1.repl1310rtf(trm(psais)); "}";
          Print #p%, " \par \cell"
          Print #p%, form1.repl1310rtf(" "); "\cell"
          Print #p%, "{\b" & talistfont & talistfontsize; "} \par "; " "; "\cell"
          Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
        Print #p%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
        Print #p%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!orchester); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.instrumentvon(l1$)
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #p%, "\par "; l1$ & iv$;
          If Len(la$) = 0 Then
            la$ = l1$ & iv$
          Else
            la$ = la$ & ", " & l1$ & iv$
          End If
        Wend
        Print #p%, "\cell"
        Print #p%, form1.repl1310rtf("" & rtmp!dirigent & ""); "\cell"
        Print #p%, "{\b" & talistfont & talistfontsize & " "; datfromsqlshort(rtmp!von); " - "; datfromsqlshort(rtmp!bis); "}\par "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "\cell"
        Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
    End If
'    End If
    rtmp.MoveNext
    Wend
    pbk% = FreeFile
    Open fn$ + ".pbk" For Input As #pbk%
    While Not EOF(pbk%)
      Line Input #pbk%, psais
      Print #p%, psais
    Wend
    Close #pbk%
    On Error Resume Next: Kill fn$ + ".pbk": On Error GoTo 0
  Else
    q% = InStr(l$, "TALISTE-K")
    If q% > 0 Then
      cmd$ = "SELECT taliste.*, talisted.talistid FROM taliste INNER JOIN talisted ON taliste.id = talisted.taid WHERE (((talisted.talistid)=""" + extsel$ + """));"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
      psais$ = ""
      While Not rtmp.EOF
    '\cellx3000\cellx6070\cellx9140
      If trm(rtmp!von) <> "" Then
      If trm(rtmp!bis) <> "" Then
        ty% = Val(Left(trm(rtmp!von), 4))
        tm% = Val(Mid(trm(rtmp!von), 6, 2))
        smcount% = tm% - ss%
        If smcount% < 0 Then
          ty% = ty% - 1
          smcount% = smcount% + sd%
        End If
        scount% = smcount% / sd%
        csais = "Saison " + trm(ty%)
        If ss% <> 1 Then csais = csais + " / " + trm(ty% + 1)
        If psais <> csais Then
          psais = csais
          Print #p%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
          Print #p%, "{\b" & talistfont & talistfontsize & " \par \cf13 "; form1.repl1310rtf(trm(psais)); "}";
          Print #p%, " \par \cell"
          Print #p%, form1.repl1310rtf(" "); "\cell"
          Print #p%, "{\b" & talistfont & talistfontsize; "} \par "; " "; "\cell"
          Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
        Print #p%, "\trowd \trgaph70\trleft-70 \cellx6111\cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
'        Print #p%, "{\b\fs24 "; form1.repl1310rtf("" & rtmp!orchester); "}\par ";
        Print #p%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!orchester); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.instrumentvon(l1$)
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #p%, "\par "; l1$ & iv$;
          If Len(la$) = 0 Then
            la$ = l1$ & iv$
          Else
            la$ = la$ & ", " & l1$ & iv$
          End If
        Wend

        Print #p%, "\cell"
        Print #p%, "{\b" & talistfont & talistfontsize & " "; datfromsqlshort(rtmp!von); " - "; datfromsqlshort(rtmp!bis); "}\par "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "\cell"
        Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
      End If
      End If
      rtmp.MoveNext
      Wend
    Else
      q% = InStr(l$, "TALISTE-OHNEDATUM")
    If q% > 0 Then
      cmd$ = "SELECT taliste.*, talisted.talistid FROM taliste INNER JOIN talisted ON taliste.id = talisted.taid WHERE (((talisted.talistid)=""" + extsel$ + """));"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
      While Not rtmp.EOF
    '\cellx3000\cellx6070\cellx9140
        If trm(rtmp!von) <> "" Then
        If trm(rtmp!bis) <> "" Then
          Print #p%, "\trowd \trgaph70\trleft-70 \cellx6111\cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
'          Print #p%, "{\b\fs24 "; form1.repl1310rtf("" & rtmp!orchester); "}\par ";
          Print #p%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!orchester); "}";
          l$ = trm(rtmp!solisten)
          la$ = ""
          While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.instrumentvon(l1$)
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #p%, "\par "; l1$ & iv$;
          If Len(la$) = 0 Then
            la$ = l1$ & iv$
          Else
            la$ = la$ & ", " & l1$ & iv$
          End If
          Wend
          Print #p%, "\cell"
'          Print #p%, "{\b\fs24    }\par "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "\cell"
          Print #p%, form1.repl1310rtf(trm(rtmp!anmerkung)); "\cell ";
          Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
       End If
        rtmp.MoveNext
        Wend
        Close ph%
      Else

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



      End If
    End If
  End If

Wend
Close #o%
Close #p%
taliste.MousePointer = 0
DoEvents
Call form1.openthisdoc(fn$, "")

engver:

fqvd$ = form1.vorlagenverzeichnis() + "\" & vorlage$ & "_engl.rtf"
If exist(fqvd$) = 0 Then
  MsgBox "Vorlage unbekannt: " + fqvd$
  Exit Sub
End If
taliste.MousePointer = 11
o% = FreeFile
Open fqvd$ For Input As #o%
p% = FreeFile
fn$ = fe$
If fn$ = "" Then
  taliste.MousePointer = 1
  Exit Sub
End If
Open fn$ For Output As #p%
On Error Resume Next
Kill fn$ + ".pbk"
On Error GoTo 0
While Not EOF(o%)
  Line Input #o%, l$
  q% = InStr(l$, "TALISTE ")
  If q% > 0 Then
    cmd$ = "SELECT taliste.*, talisted.talistid FROM taliste INNER JOIN talisted ON taliste.id = talisted.taid WHERE (((talisted.talistid)=""" + extsel$ + """)) order by von;"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
    While Not rtmp.EOF
    '\cellx3000\cellx6070\cellx9140
'    If trm(rtmp!von) <> "" Then
'    If trm(rtmp!bis) <> "" Then
    If trm(rtmp!bis) = "" Then
        csaisb = "Projects on Demand"
        pbk% = FreeFile
        Open fn$ + ".pbk" For Append As #pbk%
        If psaisb <> csaisb Then
          psaisb = csaisb
          Print #pbk%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
          Print #pbk%, "{\b" & talistfont & talistfontsize & " \par \cf13 "; form1.repl1310rtf(trm(psaisb)); "}";
          Print #pbk%, " \par \cell"
          Print #pbk%, form1.repl1310rtf(" "); "\cell"
          Print #pbk%, "{\b" & talistfont & talistfontsize; "} \par "; " "; "\cell"
          Print #pbk%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
        Print #pbk%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
        Print #pbk%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & form1.engnameof("" + rtmp!orchester + "")); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.dictionarylookup(form1.instrumentvon(l1$))
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #pbk%, "\par "; form1.engnameof("" + trm(l1$) + "") & iv$;
          If Len(la$) = 0 Then
            la$ = l1$ & iv$
          Else
            la$ = la$ & ", " & l1$ & iv$
          End If
        Wend
        Print #pbk%, "\cell"
        Print #pbk%, form1.repl1310rtf("" & form1.engnameof("" + trm(rtmp!dirigent) + "") & ""); "\cell"
        Print #pbk%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "}\cell"
        Print #pbk%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        Close #pbk%
    Else
        ty% = Val(Left(trm(rtmp!von), 4))
        tm% = Val(Mid(trm(rtmp!von), 6, 2))
        smcount% = tm% - ss%
        If smcount% < 0 Then
          ty% = ty% - 1
          smcount% = smcount% + sd%
        End If
        scount% = smcount% / sd%
        csais = "Season " + trm(ty%)
        If ss% <> 1 Then csais = csais + " / " + trm(ty% + 1)
        If psais <> csais Then
          psais = csais
          Print #p%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
          Print #p%, "{\b" & talistfont & talistfontsize & " \par \cf13 "; form1.repl1310rtf(trm(psais)); "}";
          Print #p%, " \par \cell"
          Print #p%, form1.repl1310rtf(" "); "\cell"
          Print #p%, "{\b" & talistfont & talistfontsize; "} \par "; " "; "\cell"
          Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
        End If
'        Print #p%, "\trowd \trgaph70\trleft-70 \cellx4111\cellx6804\cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
        Print #p%, "\trowd \trgaph70\trleft-70 \trbrdrt\brdrs\brdrw30 \trbrdrl\brdrs\brdrw30 \trbrdrb\brdrs\brdrw30 \trbrdrr\brdrs\brdrw30 \trbrdrh\brdrs\brdrw15 \trbrdrv\brdrs\brdrw15 \clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw30 \clbrdrb\brdrs\brdrw30 \clbrdrr \brdrs\brdrw15 \cellx4111\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw15 \cellx6804\clbrdrt\brdrs\brdrw30 \clbrdrl\brdrs\brdrw15 \clbrdrb\brdrs\brdrw30 \clbrdrr\brdrs\brdrw30 \cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
'        Print #p%, "{\b\fs24 "; form1.repl1310rtf("" & form1.engnameof("" & rtmp!orchester) & ""); "}\par ";
        Print #p%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & form1.engnameof("" & rtmp!orchester) & ""); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.dictionarylookup(form1.instrumentvon(l1$))
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #p%, "\par "; form1.engnameof("" & l1$) & iv$;

          If Len(la$) = 0 Then
            la$ = form1.engnameof("" & l1$) & iv$
          Else
            la$ = la$ & ", " & form1.engnameof("" & l1$) & iv$
          End If
        Wend
        Print #p%, "\cell"
        Print #p%, form1.repl1310rtf("" & form1.engnameof("" & rtmp!dirigent) & ""); "\cell"
        Print #p%, "{\b" & talistfont & talistfontsize & " "; datfromsqlshort(rtmp!von); " - "; datfromsqlshort(rtmp!bis); "}\par "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "\cell"
        Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
    End If
'    End If
    rtmp.MoveNext
    Wend
    pbk% = FreeFile
    Open fn$ + ".pbk" For Input As #pbk%
    While Not EOF(pbk%)
      Line Input #pbk%, psais
      Print #p%, psais
    Wend
    Close #pbk%
    On Error Resume Next: Kill fn$ + ".pbk": On Error GoTo 0
  Else
    q% = InStr(l$, "TALISTE-K")
    If q% > 0 Then
      cmd$ = "SELECT taliste.*, talisted.talistid FROM taliste INNER JOIN talisted ON taliste.id = talisted.taid WHERE (((talisted.talistid)=""" + extsel$ + """));"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
      While Not rtmp.EOF
    '\cellx3000\cellx6070\cellx9140
      If trm(rtmp!von) <> "" Then
      If trm(rtmp!bis) <> "" Then
        Print #p%, "\trowd \trgaph70\trleft-70 \cellx6111\cellx9141 \pard \intbl\tx0\tx720\tx1440\tx2160\tx2880\tx3480\tx4200\tx4920\tx5640\tx6360\tx7080\tx7800\tx8520\tx9240\tx9960 {" & talistfont & talistfontsize
'        Print #p%, "{\b\fs24 "; form1.repl1310rtf("" & form1.engnameof("" & rtmp!orchester) & ""); "}\par ";
        Print #p%, "{\b" & talistfont & talistfontsize & " "; form1.repl1310rtf("" & form1.engnameof("" & rtmp!orchester) & ""); "}";
        l$ = trm(rtmp!solisten)
        la$ = ""
        While l$ <> ""
          ps% = InStr(l$, vbCrLf)
          If ps% > 0 Then
            l1$ = Left$(l$, ps% - 1)
            l$ = Mid$(l$, ps% + 2)
          Else
            l1$ = l$
            l$ = ""
          End If
          iv$ = form1.dictionarylookup(form1.instrumentvon(l1$))
          If iv$ <> "" Then iv$ = ", " & iv$
          Print #p%, "\par "; form1.engnameof("" & l1$) & iv$;
          If Len(la$) = 0 Then
            la$ = form1.engnameof("" & l1$) & iv$
          Else
            la$ = la$ & ", " & form1.engnameof("" & l1$) & iv$
          End If
        Wend
        Print #p%, "\cell"
        Print #p%, "{\b" & talistfont & talistfontsize & " "; datfromsqlshort(rtmp!von); " - "; datfromsqlshort(rtmp!bis); "}\par "; form1.repl1310rtf("" & rtmp!anmerkung & ""); "\cell"
        Print #p%, "}\pard \intbl {" & talistfont & talistfontsize & " \row }\pard"
      End If
      End If
      rtmp.MoveNext
      Wend
    Else
'      Print #p%, l$

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






    End If
  End If

Wend
Close #o%
Close #p%

Call form1.openthisdoc(fn$, "")

taliste.MousePointer = 0
End Sub

Private Sub Command7_Click()
'd2infile = "taliste": d2insub = "Command7_Click"
Load alarmlist
Call alarmlist.settab("taliste")
alarmlist.Caption = "Angebotsliste-ID:" + Text1(0).text
End Sub

Private Sub Command8_Click()
Dim i%

'd2infile = "taliste": d2insub = "Command8_Click"
For i% = 0 To chgs.ListCount - 1
  form1.sqlqry (chgs.List(i%))
Next i%
chgs.Clear
BackColor = form1.cleancolor()
Command8.Enabled = False
End Sub

Private Sub Command9_Click()
Dim i%

'd2infile = "taliste": d2insub = "Command9_Click"
i% = List1.ListIndex
If i% >= 0 Then
  Load create2do
  Call create2do.initmsg(form1.getuserid(), form1.getuserid(), List1.List(List1.ListIndex) & " [Wiedervorlage] Angebotsliste:" + _
               List1.List(List1.ListIndex), "", Date, Left(Time, 5))
  Call create2do.SetFocus
  create2do.Text1(1).Enabled = False
  create2do.Text1(3).Enabled = False
End If

End Sub

Private Sub Form_Load()
Dim dbn$, dbn1$, s%, nflds As Integer, k%, i%, rrr

'd2infile = "taliste": d2insub = "Form_Load"
preselnodo% = 0
nochg% = 1
Randomize
Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
s% = form1.myfontsize()
List1.Font.Size = s%
Label1(1).ForeColor = form1.lnkcolor
Label1(2).ForeColor = form1.lnkcolor
Label1(3).ForeColor = form1.lnkcolor
Label1(10).ForeColor = form1.lnkcolor
Label1(11).ForeColor = form1.lnkcolor
Label1(12).ForeColor = form1.lnkcolor
Label1(19).ForeColor = form1.lnkcolor
Label1(20).ForeColor = form1.lnkcolor
Label1(21).ForeColor = form1.lnkcolor
gd2.View = lvwReport
Set colHeader = gd2.ColumnHeaders.add(, , transe("Linkbez"), 700)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Ziel"), 1200)
Set colHeader = gd2.ColumnHeaders.add(, , transe("ID"), 12)

csi% = -1
'dbpara$ = form1.getconnstr()
dbn$ = form1.getdbname()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(dbn$, dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(dbn$, False, False)
'End If
nflds = 8
For k% = 0 To 2
For i% = 0 To nflds
  On Error Resume Next
  Label1(i% + k% * (nflds + 1)).Caption = transe(form1.sqla.TableDefs("taliste").Fields(i%).name)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    Unload Me
    Exit Sub
  End If
  Text1(i% + k% * (nflds + 1)).text = ""
  Text1(i% + k% * (nflds + 1)).Font.Size = s%
Next i%
Next k%
taliste.Caption = transe("   Tournee-Angebote - AgencyProf")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command10(2).ToolTipText = transe("Zum Projekt")
Command10(1).ToolTipText = transe("Zum Projekt")
Command10(0).ToolTipText = transe("Zum Projekt")
Text1(25).ToolTipText = transe("Interne Anmerkung")
Text1(24).ToolTipText = transe("Anmerkung, die mitgedruckt wird")
Text1(21).ToolTipText = transe("scrollen, wenn nicht alle sichtbar")
Command15.Caption = transe("a&lle zeigen")
Command15.ToolTipText = transe("Markierung abwählen")
Command14.ToolTipText = transe("löschen")
Command13.ToolTipText = transe("Neue Angebotsliste")
Command12.Caption = transe("Angebotslisten      >>")
Command11.ToolTipText = transe("Schließen")
List2.ToolTipText = transe("Arten von Angebotslisten")
Command9.ToolTipText = transe("Wiedervorlage")
Command8.ToolTipText = transe("Speichern")
Command7.ToolTipText = transe("Alarm setzen")
Command6.ToolTipText = transe("Drucken")
Command4.ToolTipText = transe("Adressdetails")
Command3.ToolTipText = transe("Adressdetails")
Text1(16).ToolTipText = transe("Interne Anmerkung")
Text1(15).ToolTipText = transe("Anmerkung, die mitgedruckt wird")
Text1(14).ToolTipText = transe("Enddatum")
Text1(13).ToolTipText = transe("Anfangsdatum")
Text1(12).ToolTipText = transe("scrollen, wenn nicht alle sichtbar")
Text1(9).ToolTipText = transe("ID der Tournee")
Text1(7).ToolTipText = transe("Interne Anmerkung")
Text1(6).ToolTipText = transe("Anmerkung, die mitgedruckt wird")
Text1(5).ToolTipText = transe("Enddatum")
Text1(4).ToolTipText = transe("Anfangsdatum")
Text1(3).ToolTipText = transe("scrollen, wenn nicht alle sichtbar")
Text1(0).ToolTipText = transe("ID der Tournee")
Command2.ToolTipText = transe("Neue Tournee")
List1.ToolTipText = transe("Liste der Tourneen")
Command1.Caption = transe("&Schliessen")
Label3.Caption = transe("Tourneen")
Label3.ToolTipText = transe("Liste der Tourneen")
Label1(21).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(20).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(19).ToolTipText = transe("Doppelklick zum Auswählen")
Label2.Caption = transe("Angebotslisten")
Label1(12).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(11).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(10).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(3).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(2).ToolTipText = transe("Doppelklick zum Auswählen")
Label1(1).ToolTipText = transe("Doppelklick zum Auswählen")

Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Width = 12645
Call form1.formpos(Me)
Call rlist1
Call rlist2
Call rlist3
BackColor = form1.cleancolor()
Command14.Enabled = False
nochg% = 0

End Sub
Sub rlist1()
Dim rtmp As ADODB.Recordset, delfl As Boolean, nosel As Integer, i%
Dim selstr$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "rlist1"
List1.Clear

nosel = 1
delfl = False
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = True Then
    i% = List2.ListCount
    nosel = 0
  End If
Next i%
selstr$ = ""
If nosel = 0 Then
  cmd$ = "SELECT taliste.id,taliste.von, talisted.talistid " + _
          "FROM talisted INNER JOIN taliste ON talisted.taid = taliste.id "
  selstr$ = getsel()
Else
  cmd$ = "SELECT id,von FROM taliste "
End If
cmd$ = cmd$ + selstr$ + " order by von"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If trm(rtmp!id) <> "" Then
    If Not IsNull(rtmp!id) Then
      List1.AddItem rtmp!id & " " + transe("ab") + ": " & rtmp!von
    End If
  Else
    delfl = True
  End If
  rtmp.MoveNext
Wend
rtmp.Close
If delfl Then form1.sqlqry ("delete from taliste where id=''")
'If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub
Private Function getsel() As String
Dim s$, i%

'd2infile = "taliste": d2insub = "getsel"
s$ = ""
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = True Then
    If Len(s$) = 0 Then
      s$ = "where ((talisted.talistid='" + List2.List(i%) + "') "
    Else
      s$ = s$ + "or (talisted.talistid='" + List2.List(i%) + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"
getsel = s$


End Function

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "taliste": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld

Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub gd2_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub gd2_DblClick()
Dim r As ADODB.Recordset
Dim id$, n$, wert$, tpid$

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "gd2_DblClick"
Set lvitem = gd2.SelectedItem
On Error Resume Next
id$ = lvitem.SubItems(2)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
tpid$ = Text1(0).text: If trm(tpid$) = "" Then Exit Sub
'MsgBox "ID=" & id$
c$ = "select * from auftritthigru where id='" & id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  wert$ = trm(Mid(r!feldname, 6)) + "=" + trm(r!felddaten)
  wert$ = InputBox(transe("Link ändern:"), transe("Link festlegen"), wert$)
  If wert$ = "" Then Exit Sub
  p% = InStr(wert$, "=")
  If p% = 0 Then wert$ = trm(Mid(r!feldname, 6)) + "=" + wert$
  p% = InStr(wert$, "=")
  If p% > 0 Then
    n$ = trm(Mid$(wert$, p% + 1))
    wert$ = trm(Left$(wert$, p% - 1))
    cmd$ = "delete from auftritthigru where id='" + tpid$ + " " + wert$ + "' and auftrittstyp='taliste'"
    Call form1.sqlqry(cmd$)
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values('"
    cmd$ = cmd$ + tpid$ + " " + wert$ + "','" + tpid$ + "','taliste','link_" + wert$ + "','" + n$ + "')"
    Call form1.sqlqry(cmd$)
    Call rgd2
  Else
    MsgBox ("Syntaxfehler: Linkbezeichnung=Linkziel")
  End If
End If
End Sub

Private Sub Label1_DblClick(Index As Integer)
Dim nfls As Integer, wi%, s$

'd2infile = "taliste": d2insub = "Label1_DblClick"
nfls = 8

wi% = Index Mod (nfls + 1)
Select Case wi%
  Case 1: s$ = transe("Orchester")
  Case 2: s$ = transe("Dirigent")
  Case 3: s$ = transe("Künstler")
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
    If wi% <> 3 Or Text1(Index).text = "" Then
      Text1(Index).text = form1.getnamebyid(adrselect.sel_getselected())
    Else
      Text1(Index).text = Text1(Index).text + Chr$(13) + Chr$(10) + form1.getnamebyid(adrselect.sel_getselected())
    End If
  End If
  Unload adrselect
End If

End Sub

Private Sub List1_Click()
Dim rtmp As ADODB.Recordset, nflds As Integer, i%, k%, li%, id$, p%, id0$

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "List1_Click"
Call savecheck
nflds = 8

i% = 0

For k% = 0 To 2
For i% = 0 To nflds
  Label1(i% + k% * (nflds + 1)).Caption = transe(form1.sqla.TableDefs("taliste").Fields(i%).name)
  Text1(i% + k% * (nflds + 1)).text = ""
Next i%
Next k%

li% = List1.ListIndex
csi% = li%
For k% = 0 To 2
If li% + k% < List1.ListCount Then

id$ = List1.List(li% + k%)
p% = InStr(id$, " " + transe("ab") + ": ")
If p% > 0 Then id$ = Left$(id$, p% - 1)
If id$ = "" Then Exit Sub
If k% = 0 Then id0$ = id$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT * FROM taliste where id='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
If Not rtmp.EOF Then
  For i% = 0 To nflds
    If Not IsNull(rtmp.Fields(i%)) Then
      If i% < 4 Or i% > 5 Then
        Text1(i% + k% * (nflds + 1)).text = rtmp.Fields(i%)
      Else
        Text1(i% + k% * (nflds + 1)).text = datfromsql(rtmp.Fields(i%))
      End If
    End If
  Next i%
End If

End If
Next k%
For i% = 0 To List3.ListCount - 1
  preselnodo% = 1: List3.Selected(i%) = False
Next i%
For i% = 0 To List2.ListCount - 1
  preselnodo% = 1: List2.Selected(i%) = False
Next i%
preselnodo% = 0
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT * FROM talisted where taid='" + id0$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
While Not rtmp.EOF
  For k% = 0 To List2.ListCount - 1
    If List2.List(k%) = rtmp!talistid Then
      preselnodo% = 1: List2.Selected(k%) = True
      k% = List2.ListCount
    End If
  Next k%
  rtmp.MoveNext
Wend
If Not form1.isfieldmissing("opt_talisted1", "id") Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rtmp.Open "SELECT talistid FROM opt_talisted1 where taid='" + id0$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
  While Not rtmp.EOF
    For k% = 0 To List3.ListCount - 1
      If List3.List(k%) = rtmp!talistid Then
        preselnodo% = 1: List3.Selected(k%) = True
        k% = List3.ListCount
      End If
    Next k%
    rtmp.MoveNext
  Wend
End If
Call rgd2
preselnodo% = 0
BackColor = form1.cleancolor()

End Sub


Private Sub List1_DblClick()
Dim r As ADODB.Recordset, s As ADODB.Recordset, p%
Dim extsel$, dfl%, solo$, msolo$, id$, hp$

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "List1_DblClick"
Call Command8_Click
id$ = List1.List(List1.ListIndex)
p% = InStr(id$, " ab: ")
If p% > 0 Then id$ = Left$(id$, p% - 1)
If id$ = "" Then Exit Sub

Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open "SELECT * FROM taliste where id='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open "SELECT id FROM tplan where id='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
  If r.EOF Then
    p% = 0
    While p% < List2.ListCount
      If List2.Selected(p%) Then
        extsel$ = List2.List(p%): dfl% = 0
        Select Case LCase$(extsel$)
          Case "künstler": dfl% = 1
          Case "orchester": dfl% = 1
          Case "kammermusik": dfl% = 1
          Case "crossover": dfl% = 1
          Case Else:
        End Select
        If dfl% = 1 Then
          hp$ = extsel$
          p% = List2.ListCount
        End If
      End If
      p% = p% + 1
    Wend
    If trm(s!von & "") = "" Then
      MsgBox ("Startdatum fehlt. Ein Tourneeplan kann nicht angelegt werden.")
      Exit Sub
    End If
    If trm(s!bis & "") = "" Then
      MsgBox ("Enddatum fehlt. Ein Tourneeplan kann nicht angelegt werden.")
      Exit Sub
    End If
    If Not IsNull(s!solisten) Then
      p% = InStr(s!solisten, Chr$(13) + Chr$(10))
      If p% > 0 Then
        solo$ = form1.getidbyname(Left$(s!solisten, p% - 1))
        msolo$ = Mid$(s!solisten, p% + 2)
      Else
        solo$ = form1.getidbyname(s!solisten)
        msolo$ = ""
      End If
    End If
    form1.sqlqry ("insert into tplan (id) values('" & id$ & "')")
    form1.sqlqry ("update tplan set orchester='" & form1.getidbyname("" & s!orchester & "") & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set solist='" & form1.getidbyname("" & solo$ & "") & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set dirigent='" & form1.getidbyname("" & s!dirigent & "") & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set mehr_solisten='" & form1.getidbyname("" & msolo$ & "") & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set von='" & s!von & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set bis='" & s!bis & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set anmerkung='" & trm(s!anmerkung) & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set anmerkungintern='" & trm(s!intern) & "' where id='" & id$ & "'")
    form1.sqlqry ("update tplan set hauptperson='" & hp$ & "' where id='" & id$ & "'")
  End If
End If
Load tplan
Call tplan.rlist1
Call tplan.gotorec(id$)
Call tplan.SetFocus
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, sq$, ask%

'd2infile = "taliste": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  idx% = csi%
  If idx% < 0 Then Exit Sub
  csi% = -1
  ask% = MsgBox("Wirklich löschen?", vbYesNo + vbCritical + vbDefaultButton2, "Angebot " & List1.List(csi%) & " löschen?")
  If ask% = vbYes Then
    id$ = List1.List(idx%)
    id$ = Left$(id$, InStr(id$, " " + transe("ab") + ": ") - 1)
    If id$ = "" Then Exit Sub
    sq$ = "delete from talisted where taid='" + id$ + "'"
    Call form1.sqlqry(sq$)
    sq$ = "delete from taliste where id='" + id$ + "'"
    Call form1.sqlqry(sq$)
  End If
End If
Call rlist1

End Sub

Private Sub List2_Click()
Dim i%, j%, nid$

'd2infile = "taliste": d2insub = "List2_Click"
If preselnodo% = 1 Then
  preselnodo% = 0
  Exit Sub
End If
i% = List1.ListIndex
If i% >= 0 Then
  j% = List2.ListIndex
  If List2.Selected(j%) = False Then
    Call form1.sqlqry("delete from talisted where talistid='" + List2.List(j%) + "' and taid='" + Text1(0).text + "'")
  Else
    Call form1.sqlqry("delete from talisted where talistid='" + List2.List(j%) + "' and taid='" + Text1(0).text + "'")
    nid$ = form1.newid("talisted", "id", 18)
    Call form1.sqlqry("insert into talisted (id,talistid,taid) values('" + nid$ + "','" + List2.List(j%) + "','" + Text1(0).text + "')")
  End If
Else
  Call rlist1
End If
End Sub

Private Sub List3_Click()
Dim i%, j%, nid$

If form1.isfieldmissing("opt_talisted1", "id") Then Exit Sub
If preselnodo% = 1 Then
  preselnodo% = 0
  Exit Sub
End If
i% = List1.ListIndex
If i% >= 0 Then
  j% = List3.ListIndex
  If List3.Selected(j%) = False Then
    Call form1.sqlqry("delete from opt_talisted1 where talistid='" + List3.List(j%) + "' and taid='" + Text1(0).text + "'")
  Else
    Call form1.sqlqry("delete from opt_talisted1 where talistid='" + List3.List(j%) + "' and taid='" + Text1(0).text + "'")
    nid$ = form1.newid("opt_talisted1", "id", 18)
    Call form1.sqlqry("insert into opt_talisted1 (id,talistid,taid) values('" + nid$ + "','" + List3.List(j%) + "','" + Text1(0).text + "')")
  End If
Else
  Call rlist1
End If

End Sub

Private Sub Text1_Change(Index As Integer)
'd2infile = "taliste": d2insub = "Text1_Change"
If nochg% = 1 Then Exit Sub
BackColor = form1.dirtycolor()
Command8.Enabled = True
End Sub

Private Sub Text1_DblClick(Index As Integer)

'd2infile = "taliste": d2insub = "Text1_DblClick"
If Index = 4 Or Index = 5 Then
  With frmCalendar
    .init Text1(Index), Text1(Index).text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1(Index).text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)

'd2infile = "taliste": d2insub = "Text1_GotFocus"
prvw$ = Trim$(Text1(Index).text)
'is ja eigentlich peinlich, aber ...
If Index = 4 Or Index = 5 Then prvw$ = "huhu"
End Sub

Public Sub Text1_LostFocus(Index As Integer)
Dim nflds As Integer, rtmp As QueryDef, r As ADODB.Recordset, chg2$, c$
Dim i%, k%, id$, f$, w$, prv$, antw As Integer, ask%

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "Text1_LostFocus"
nflds = 8

i% = Index Mod (nflds + 1)
k% = Int(Index / (nflds + 1))
id$ = Text1(k% * (nflds + 1)).text
If id$ = "" Then
  Text1(Index).text = ""
  Exit Sub
End If
f$ = Label1(Index).Caption
w$ = trm(Text1(Index).text)
If w$ <> prvw$ Then
  If i% = 0 Then
    antw = MsgBox("Angebot " + prv$ + " umbenennen?", vbYesNo + vbCritical + vbDefaultButton2, "Umbenennen? (nicht empfehlenswert)")
    If antw <> vbYes Then
      BackColor = form1.cleancolor()
      Command8.Enabled = False
      Text1(i%).text = prvw$
      BackColor = form1.cleancolor()
      Command8.Enabled = False
      Exit Sub
    End If
  End If
  chg2$ = ""
  If i% = 4 Or i% = 5 Then
    w$ = datum2sql(w$)
    prvw$ = "huhu"
  End If
  c$ = "select id from tplan where id='" + id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, dbOpenDynaset, dbReadOnly
  If Not r.EOF Then
    ask% = vbYes
    If form1.getusersetting("AngebotProjektDatumPrüfen") = "nein" Then ask% = vbNo
    If ask% = vbYes Then
      chg2$ = "update tplan set " + f$ & "='" + w$ & "' where id='" + id$ & "'"
    End If
  End If
  w$ = "'" & w$ & "'": If w$ = "''" Then w$ = "NULL"
  BackColor = form1.dirtycolor()
  Command8.Enabled = True
  Text1(8).text = datum2sql(Date) & " " & Time
  If f$ <> "id" Then
    chgs.AddItem "update taliste set " + transo(f$) + "=" + w$ + " where id='" + id$ + "'"
    If chg2$ <> "" Then chgs.AddItem chg2$
  Else
    chgs.AddItem "update taliste set " + transo(f$) + "=" + w$ + " where id='" + prvw$ + "'"
  End If
End If

End Sub

Sub savecheck()
Dim antw As Integer

'd2infile = "taliste": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command8_Click
  End If
End If
BackColor = form1.cleancolor()
End Sub

Sub rlist2()
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "rlist2"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT id FROM talisttyp order by id", form1.adoc, dbOpenDynaset, dbReadOnly

List2.Clear
While Not rtmp.EOF
  If Not IsNull(rtmp!id) Then List2.AddItem rtmp!id
  rtmp.MoveNext
Wend
End Sub

Sub rlist3()
Dim rtmp As ADODB.Recordset, o$

o$ = "SELECT wert FROM sysvars where instr(owner,'sysvar_system_tatyp')>0 order by wert"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open o$, form1.adoc, adOpenDynamic, adLockReadOnly

List3.Clear
While Not rtmp.EOF
  If Not IsNull(rtmp!wert) Then List3.AddItem trm(rtmp!wert)
  rtmp.MoveNext
Wend
End Sub

Public Sub presel(t$)
Dim i%

'd2infile = "taliste": d2insub = "presel"
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = False Then
    If List2.List(i%) = t$ Then
      List2.Selected(i%) = True
      Exit Sub
    End If
  End If
Next i%

End Sub

Sub rgd2()
Dim r As ADODB.Recordset, c$, lvlitem As ListItem, lnbez As String, tpid$

Dim d2infile As String, d2insub As String
d2infile = "taliste": d2insub = "rgd2"

gd2.ListItems.Clear
tpid$ = Text1(0).text: If trm(tpid$) = "" Then Exit Sub
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM auftritthigru where auftrittsid='" + tpid$ & "' and auftrittstyp='taliste' and instr(feldname,'link_')=1 order by id", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName) values('"
  cmd$ = cmd$ + tpid$ + " Bio','" + tpid$ + "','taliste','link_Bio')"
  Call form1.sqlqry(cmd$)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName) values('"
  cmd$ = cmd$ + tpid$ + " Foto','" + tpid$ + "','taliste','link_Foto')"
  Call form1.sqlqry(cmd$)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName) values('"
  cmd$ = cmd$ + tpid$ + " Detail','" + tpid$ + "','taliste','link_Detail')"
  Call form1.sqlqry(cmd$)
  r.Close
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, "SELECT * FROM auftritthigru where auftrittsid='" + tpid$ & "' and auftrittstyp='taliste' and instr(feldname,'link_')=1 order by id", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
End If
While Not r.EOF
  lnbez = trm(Mid(trm(r!feldname), 6))
  wer$ = trm(r!felddaten)
  Set lvitem = gd2.ListItems.add(, , lnbez)
  lvitem.SubItems(1) = wer$
  lvitem.SubItems(2) = r!id
  r.MoveNext
Wend

End Sub

