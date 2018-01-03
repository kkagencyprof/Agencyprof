VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form besetzung 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Besetzungen"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   5925
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows-Standard
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
      Left            =   480
      TabIndex        =   75
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox dflt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Standard"
      Height          =   255
      Left            =   6600
      TabIndex        =   74
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "als neu"
      Height          =   375
      Left            =   720
      TabIndex        =   73
      Top             =   5400
      Width           =   615
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
      Left            =   840
      Picture         =   "besetzung.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   72
      ToolTipText     =   "Neue Besetzung"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "besetzung.frx":0392
      Style           =   1  'Grafisch
      TabIndex        =   70
      ToolTipText     =   "Speichern"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox bes 
      Height          =   1245
      Index           =   29
      Left            =   1320
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   4560
      Width           =   6255
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   28
      Left            =   2160
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   27
      Left            =   1680
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   26
      Left            =   1200
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   25
      Left            =   720
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   24
      Left            =   240
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   23
      Left            =   5640
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox bes 
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
      Index           =   22
      Left            =   4680
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   21
      Left            =   3840
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   20
      Left            =   3000
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   19
      Left            =   2160
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   1320
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   17
      Left            =   3840
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   16
      Left            =   2160
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   15
      Left            =   1680
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   3000
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   2160
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   12
      Left            =   1320
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   11
      Left            =   4680
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   3840
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   8
      Left            =   3000
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   7
      Left            =   2160
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   6
      Left            =   1320
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   5
      Left            =   4680
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox bes 
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   3840
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   2
      Left            =   3000
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Index           =   1
      Left            =   2160
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox bes 
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
      Left            =   1320
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1920
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1080
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      Picture         =   "besetzung.frx":0739
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   5400
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label bid 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3000
      TabIndex        =   71
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
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
      Index           =   34
      Left            =   120
      TabIndex        =   39
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   33
      Left            =   720
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
      Height          =   255
      Index           =   32
      Left            =   1560
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   31
      Left            =   2400
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   30
      Left            =   3240
      TabIndex        =   35
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   29
      Left            =   4080
      TabIndex        =   34
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   28
      Left            =   2880
      TabIndex        =   33
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   27
      Left            =   5640
      TabIndex        =   32
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   26
      Left            =   4680
      TabIndex        =   31
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   25
      Left            =   3840
      TabIndex        =   30
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   24
      Left            =   3000
      TabIndex        =   29
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
      Height          =   255
      Index           =   23
      Left            =   2160
      TabIndex        =   28
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   22
      Left            =   1320
      TabIndex        =   27
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
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
      Index           =   21
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   20
      Left            =   3840
      TabIndex        =   25
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   19
      Left            =   2880
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   18
      Left            =   720
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
      Height          =   255
      Index           =   16
      Left            =   2160
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
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
      Index           =   14
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   13
      Left            =   4680
      TabIndex        =   18
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   12
      Left            =   2160
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   15
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
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
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Klarinette"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flöte"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Besetzungen:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label werknam 
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label werkid 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "besetzung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ins$(4, 6)
Dim dbl$(4, 1 To 6)

Private Sub bes_Change(Index As Integer)

'd2infile = "besetzung": d2insub = "bes_Change"
If InStr(bes(Index).text, "/") > 0 Then
  t$ = strrepl(bes(Index).text, "/", "-")
  bes(Index).text = t$
End If
Me.BackColor = form1.dirtycolor()
Command4.Enabled = True

End Sub



Private Sub Command1_Click()
'd2infile = "besetzung": d2insub = "Command1_Click"
Unload besetzung

End Sub

Private Sub Command11_Click()
Dim wid$

'd2infile = "besetzung": d2insub = "Command11_Click"
wid$ = werknam.Caption
Call savecheck

bid.Caption = ""
Call Command4_Click
Call rlist1

End Sub

Private Sub Command19_Click()
'd2infile = "besetzung": d2insub = "Command19_Click"
Call form1.handbuchcall("07-Werkeverzeichnis.htm")

End Sub

Private Sub Command2_Click()

'd2infile = "besetzung": d2insub = "Command2_Click"
bid.Caption = ""
Call Command4_Click
Call rlist1

End Sub

Private Sub Command4_Click()

'd2infile = "besetzung": d2insub = "Command4_Click"
If werkid.Caption = "" Then Exit Sub
If bid.Caption = "" Then
  bid.Caption = form1.newid("b_loc", "id", 28)
  c$ = "insert into b_loc (id,wid) values('" & bid.Caption & "','" & werkid.Caption & "')"
  Call form1.sqlqry(c$)
End If

For i% = 0 To 4
  For j% = 1 To 6
    If dbl$(i%, j%) <> "" Then
      k% = i% * 6 + (j% - 1)
      If trm(bes(k%).text) = "" Then bes(k%).text = "-"
      DoEvents
      c$ = "update b_loc set " & LCase(dbl$(i%, j%)) & "='" & bes(k%).text & "' where (id='" & bid.Caption & "')"
      Call form1.sqlqry(c$)
      bes(k%).text = ""
      DoEvents
    End If
  Next j%
Next i%
If dflt.value > 0 Then
  c$ = "update b_loc set dflt=0 where (wid='" & werkid.Caption & "')"
  Call form1.sqlqry(c$)
  c$ = "update b_loc set dflt=1 where (id='" & bid.Caption & "')"
  Call form1.sqlqry(c$)
Else
  c$ = "update b_loc set dflt=0 where (id='" & bid.Caption & "')"
  Call form1.sqlqry(c$)
End If

BackColor = form1.cleancolor()
Call rlist1
BackColor = form1.cleancolor()
Call reload(bid.Caption)
BackColor = form1.cleancolor()

End Sub

Private Sub dflt_Click()
'd2infile = "besetzung": d2insub = "dflt_Click"
Me.BackColor = form1.dirtycolor()
Command4.Enabled = True

End Sub

Private Sub Form_Load()

'd2infile = "besetzung": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Randomize


s% = form1.myfontsize()
List1.Font.Size = s%
werknam.Font.Size = s%
For i% = 0 To 29: bes(i%).Font.Size = s%: Next i%
For i% = 0 To 4: For j% = 0 To 6: ins$(i%, j%) = "": Next j%: Next i%

ins$(0, 0) = "Holz"
ins$(0, 1) = "Flöte"
ins$(0, 2) = "Oboe"
ins$(0, 3) = "Klarinette"
ins$(0, 4) = "Fagott"
ins$(0, 6) = "Andere"

ins$(1, 0) = "Blech"
ins$(1, 1) = "Horn"
ins$(1, 2) = "Trompete"
ins$(1, 3) = "Posaune"
ins$(1, 4) = "Tuba"
ins$(1, 6) = "Andere"

ins$(2, 0) = "Schlagwerk"
ins$(2, 1) = "Pauke"
ins$(2, 2) = "Schlagwerk"
ins$(2, 3) = "-------"
ins$(2, 6) = "Andere"

ins$(3, 0) = "Streicher"
ins$(3, 1) = "1. Violine"
ins$(3, 2) = "2. Violine"
ins$(3, 3) = "Viola"
ins$(3, 4) = "Cello"
ins$(3, 5) = "Kontrabass"
ins$(3, 6) = "Andere"

ins$(4, 0) = "Andere"
ins$(4, 6) = "Andere"

dbl$(0, 1) = "floete"
dbl$(0, 2) = "Oboe"
dbl$(0, 3) = "Klarinette"
dbl$(0, 4) = "Fagott"
dbl$(0, 6) = "holz_sonst"

dbl$(1, 1) = "Horn"
dbl$(1, 2) = "Trompete"
dbl$(1, 3) = "Posaune"
dbl$(1, 4) = "Tuba"
dbl$(1, 6) = "blech_sonst"

dbl$(2, 1) = "Pauke"
dbl$(2, 2) = "Triangel"
dbl$(2, 3) = "Becken"
dbl$(2, 6) = "schlagwerk_sonst"

dbl$(3, 1) = "Violine1"
dbl$(3, 2) = "Violine2"
dbl$(3, 3) = "Viola"
dbl$(3, 4) = "Cello"
dbl$(3, 5) = "Kontrabass"
dbl$(3, 6) = "streicher_sonst"

dbl$(4, 6) = "sonst"

'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

'dbpara$ = form1.getconnstr()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
'End If
For i% = 0 To 4
  For j% = 0 To 6
    k% = i% * 7 + j%
    Label2(k%).Caption = ins$(i%, j%)
  Next j%
Next i%

Call rlist1
If List1.ListIndex < 0 Then
  Call reload("NULL")
Else
  Call reload(bid.Caption)
End If
Show

End Sub

Private Sub Form_Resize()
'd2infile = "besetzung": d2insub = "Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

'd2infile = "besetzung": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub


Private Sub List1_Click()
Dim id$

'd2infile = "besetzung": d2insub = "List1_Click"
Call savecheck
id$ = List1.List(List1.ListIndex)
If InStr(id$, "ID:") = 0 Then Exit Sub
id$ = Mid$(id$, InStr(id$, "ID:") + 3)
bid.Caption = id$
Call reload(id$)

End Sub

Private Sub List1_DblClick()
Dim i As Integer, rrr, fnam As String, p As Integer, bnam As String

'd2infile = "besetzung": d2insub = "List1_DblClick"
li = List1.ListIndex
If li < 0 Then Exit Sub
bnam = List1.List(li)
p = InStr(bnam, "ID:") - 1
If p > 1 Then bnam = trm(Left(bnam, p))
If Left(werkid.Caption, 2) = "T:" Then
  i = 0
  Do
    On Error Resume Next
    fnam = auftritt.Label2(i).Caption
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      If fnam = word1(werknam.Caption) Then
        Call auftritt.Text2_GotFocus(i)
        auftritt.Text2(i).text = bnam
        Call auftritt.Text2_LostFocus(i)
        Call Command1_Click
        Exit Sub
      End If
    End If
    i = i + 1
  Loop Until i > 33 Or auftritt.Label2(i).Visible = False Or rrr
End If
If werkvz.isviz Then
  If besetzung.werkid.Caption = werkvz.Text4(0).text Then
    If InStr(bnam, "Standard:") = 1 Then bnam = trm(Mid(bnam, 10))
    werkvz.Text4(4).text = bnam
  End If
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "besetzung": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  id$ = List1.List(List1.ListIndex)
  If InStr(id$, "ID:") = 0 Then Exit Sub
  id$ = Mid$(id$, InStr(id$, "ID:") + 3)
  c$ = "SELECT * FROM programmliste where besetztid='" & id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    MsgBox "Diese Besetzung kommt mindestens im Programm:" & vbCrLf & r!programmid & " vor. Löschen nicht möglich."
    Exit Sub
  End If
  ask% = MsgBox("Wirklich löschen?" & vbCrLf & form1.bestzstr(id$), vbYesNo + vbCritical + vbDefaultButton2, "Besetzung löschen?")
  If ask% = vbYes Then
    If id$ = "" Then Exit Sub
    sq$ = "delete from b_loc where id='" + id$ + "'"
    Call form1.sqlqry(sq$)
  End If
  Call rlist1
End If

End Sub

Private Sub werkid_Change()

'd2infile = "besetzung": d2insub = "werkid_Change"
werknam.Caption = form1.getwerknamebyid(werkid.Caption)
Call rlist1
Call reload("NULL")

End Sub

Private Sub werknam_Change()
'd2infile = "besetzung": d2insub = "werknam_Change"
If Left(werkid.Caption, 2) = "T:" Then
  If List1.ListCount > 0 Then List1.ListIndex = 0
End If
End Sub

Private Sub werknam_DblClick()
'd2infile = "besetzung": d2insub = "werknam_DblClick"
werkvz.Text2.text = werknam.Caption
Call werkvz.Text2_Change
End Sub

Private Sub savecheck()
'd2infile = "besetzung": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
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

Sub rlist1()
Dim rtmp As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "besetzung": d2insub = "rlist1"
c$ = "SELECT * FROM b_loc where wid='" & werkid.Caption & "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
List1.Clear
While Not rtmp.EOF
  ad$ = form1.bestzstr(rtmp!id)
  List1.AddItem ad$ & Space$(120) & "ID:" & rtmp!id
  rtmp.MoveNext
Wend
For i% = 0 To List1.ListCount - 1
  If InStr(List1.List(i%), "Standard") = 1 Then
    List1.ListIndex = i%
    Call List1_Click
    DoEvents
    Exit For
  End If
Next i%

End Sub

Sub reload(arg)
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "besetzung": d2insub = "reload"
MousePointer = 11
For i% = 0 To 4
  For j% = 1 To 6
    k% = i% * 6 + (j% - 1)
    bes(k%).text = ""
  Next j%
Next i%
dflt.value = 0
If arg = "NULL" Then bid.Caption = ""
c$ = "SELECT * FROM b_loc where id='" & arg & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  p% = 0
  For i% = 0 To 4
    For j% = 1 To 6
      k% = i% * 6 + (j% - 1)
      If dbl$(i%, j%) <> "" Then
        bes(k%).text = "" & trm(r.Fields(p% + 2).value)
        p% = p% + 1
      End If
    Next j%
  Next i%
  If Not IsNull(r!dflt) Then dflt.value = r!dflt
End If
MousePointer = 0
Me.BackColor = form1.cleancolor()
Command4.Enabled = False

End Sub
