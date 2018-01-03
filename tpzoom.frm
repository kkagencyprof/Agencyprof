VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form tpzoom 
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   7680
      Picture         =   "tpzoom.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   113
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   7080
      Picture         =   "tpzoom.frx":0B82
      Style           =   1  'Grafisch
      TabIndex        =   112
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   111
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   110
      Top             =   5880
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   109
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   108
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   107
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton vsw 
      Caption         =   "?"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   106
      Top             =   7080
      Width           =   255
   End
   Begin VB.CommandButton vsw 
      Caption         =   "?"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   105
      Top             =   5520
      Width           =   255
   End
   Begin VB.CommandButton vsw 
      Caption         =   "?"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   104
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton vsw 
      Caption         =   "?"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   103
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton vsw 
      Caption         =   "?"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   102
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton hsw 
      Caption         =   "?"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   101
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton hsw 
      Caption         =   "?"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   100
      Top             =   5160
      Width           =   255
   End
   Begin VB.CommandButton hsw 
      Caption         =   "?"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   99
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton hsw 
      Caption         =   "?"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   98
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton hsw 
      Caption         =   "?"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   97
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      Picture         =   "tpzoom.frx":1704
      Style           =   1  'Grafisch
      TabIndex        =   96
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox beg 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   94
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox beg 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   92
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox beg 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   90
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox beg 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   88
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox beg 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   86
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox styp 
      Height          =   315
      Index           =   4
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   85
      Text            =   "styp"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.ComboBox styp 
      Height          =   315
      Index           =   3
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   84
      Text            =   "styp"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ComboBox styp 
      Height          =   315
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   83
      Text            =   "styp"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox styp 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   82
      Text            =   "styp"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox styp 
      Height          =   315
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   81
      Text            =   "styp"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton shauf 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "tpzoom.frx":226A
      Style           =   1  'Grafisch
      TabIndex        =   75
      ToolTipText     =   "Zeige mehr Details"
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton shauf 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "tpzoom.frx":33E4
      Style           =   1  'Grafisch
      TabIndex        =   74
      ToolTipText     =   "Zeige mehr Details"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton shauf 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "tpzoom.frx":455E
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Zeige mehr Details"
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton shauf 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "tpzoom.frx":56D8
      Style           =   1  'Grafisch
      TabIndex        =   72
      ToolTipText     =   "Zeige mehr Details"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton shauf 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "tpzoom.frx":6852
      Style           =   1  'Grafisch
      TabIndex        =   71
      ToolTipText     =   "Zeige mehr Details"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox hinw 
      Height          =   1335
      Index           =   1
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "tpzoom.frx":79CC
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox hinw 
      Height          =   1335
      Index           =   0
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   56
      Text            =   "tpzoom.frx":79D2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox hinw 
      Height          =   1335
      Index           =   2
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   55
      Text            =   "tpzoom.frx":79D8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.ComboBox prg 
      Height          =   315
      Index           =   4
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   50
      Text            =   "Combo1"
      Top             =   7440
      Width           =   2055
   End
   Begin VB.ComboBox prg 
      Height          =   315
      Index           =   3
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   49
      Text            =   "Combo1"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.ComboBox prg 
      Height          =   315
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   48
      Text            =   "Combo1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ComboBox prg 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   46
      Text            =   "Combo1"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox prg 
      Height          =   315
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   45
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox abez 
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
      Index           =   4
      Left            =   2760
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox abez 
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
      Left            =   2760
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox abez 
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
      Left            =   2760
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox abez 
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
      Left            =   2760
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox abez 
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
      Left            =   2760
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox hinw 
      Height          =   1335
      Index           =   4
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "tpzoom.frx":79DE
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox hinw 
      Height          =   1335
      Index           =   3
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "tpzoom.frx":79E4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox hon 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox hon 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox hon 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox hon 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox hon 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox veran 
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox veran 
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox veran 
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox veran 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox veran 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox hlle 
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox hlle 
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox hlle 
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox hlle 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox hlle 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox aort 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox aort 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox aort 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox aort 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox aort 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   7680
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tpzoom.frx":79EA
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   7920
      Width           =   1695
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3120
      Top             =   7800
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   95
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   93
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   91
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   89
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   87
      Top             =   120
      Width           =   135
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   8160
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   8160
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   8160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
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
      Index           =   4
      Left            =   840
      TabIndex        =   80
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
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
      Index           =   3
      Left            =   840
      TabIndex        =   79
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
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
      Index           =   2
      Left            =   840
      TabIndex        =   78
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
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
      Index           =   1
      Left            =   840
      TabIndex        =   77
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
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
      Left            =   840
      TabIndex        =   76
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label aid 
      Caption         =   "Label10"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   70
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label aid 
      Caption         =   "Label10"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   69
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label aid 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   68
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label aid 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label aid 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   66
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lhlle 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   65
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label lveran 
      BackStyle       =   0  'Transparent
      Caption         =   "Veranstalter"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   64
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lhlle 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   63
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lveran 
      BackStyle       =   0  'Transparent
      Caption         =   "Veranstalter"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   62
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lhlle 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   61
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lveran 
      BackStyle       =   0  'Transparent
      Caption         =   "Veranstalter"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   60
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lhlle 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   59
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lveran 
      BackStyle       =   0  'Transparent
      Caption         =   "Veranstalter"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   58
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm"
      Height          =   255
      Left            =   2520
      TabIndex        =   54
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm"
      Height          =   255
      Left            =   2520
      TabIndex        =   53
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm"
      Height          =   255
      Left            =   2520
      TabIndex        =   52
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm"
      Height          =   255
      Left            =   2520
      TabIndex        =   51
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm"
      Height          =   255
      Left            =   2520
      TabIndex        =   47
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lveran 
      BackStyle       =   0  'Transparent
      Caption         =   "Veranstalter"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   44
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lhlle 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   43
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorar"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorar"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorar"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorar"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorar"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label adatum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label adatum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label adatum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label adatum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label adatum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "tpzoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id0$
Dim nochg%
Dim fl_abez%(5)
Dim fl_ort%(5), fl_hon%(5), fl_hlle%(5), fl_prg%(5), fl_hinw%(5), fl_veran%(5), fl_beg%(5)
Dim tmrbusy%, tpzmode%

Private Sub abez_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "abez_Change"
  If nochg% = 1 Then Exit Sub
  fl_abez%(Index) = 1
End Sub

Private Sub abez_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$

'd2infile = "tpzoom": d2insub = "abez_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_abez%(n%) = 0 Then Exit Sub
fl_abez%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update auftritt set bezeichnung='" & abez(n%).text & "' where id='" + a_id$ + "'"

MsgBox cmd$

End Sub


Private Sub aort_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "aort_Change"
  If nochg% = 1 Then Exit Sub
  fl_ort%(Index) = 1
End Sub

Private Sub aort_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$

'd2infile = "tpzoom": d2insub = "aort_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_ort%(n%) = 0 Then Exit Sub
fl_ort%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update auftritt set ort='" & aort(n%).text & "' where id='" + a_id$ + "'"

Call form1.sqlqry(cmd$)


End Sub

Private Sub beg_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "beg_Change"
  If nochg% = 1 Then Exit Sub
  fl_beg%(Index) = 1

End Sub

Private Sub beg_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$

'd2infile = "tpzoom": d2insub = "beg_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_beg%(n%) = 0 Then Exit Sub
fl_beg%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update auftritt set zeit='" & Left(beg(n%).text, 5) & "' where id='" & a_id$ & "'"

Call form1.sqlqry(cmd$)


End Sub

Private Sub Command1_Click()

'd2infile = "tpzoom": d2insub = "Command1_Click"
Unload tpzoom

End Sub

Private Sub Command3_Click()
'd2infile = "tpzoom": d2insub = "Command3_Click"
id0$ = ""
Call Timer1_Timer
End Sub


Private Sub Command5_Click(Index As Integer)
Dim c$

'd2infile = "tpzoom": d2insub = "Command5_Click"
c$ = trm(prg(Index).text)
If c$ <> "" Then
  If Len(c$) > 0 Then
    Load prog
    Call prog.SetFocus
    Call prog.selectone(c$)
  End If
End If

End Sub

Private Sub Command6_Click()
Dim i%

'd2infile = "tpzoom": d2insub = "Command6_Click"
i% = tplan.List6.ListIndex + 1
If i% >= tplan.List6.ListCount Then i% = tplan.List6.ListCount - 1
tplan.List6.ListIndex = i%
DoEvents
Call Timer1_Timer

End Sub

Private Sub Command7_Click()
Dim i%

'd2infile = "tpzoom": d2insub = "Command7_Click"
i% = tplan.List6.ListIndex - 1
If i% < 0 Then i% = 0
tplan.List6.ListIndex = i%
DoEvents
Call Timer1_Timer

End Sub

Private Sub Form_Load()
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "Form_Load"
axsResizer1.SaveControlPositions

tmrbusy% = 0
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
BackColor = form1.cleancolor()
id0$ = ""
tpzmode% = 0

End Sub

Private Sub Form_Resize()
'd2infile = "tpzoom": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "tpzoom": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub

Private Sub hinw_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "hinw_Change"
  If nochg% = 1 Then Exit Sub
  fl_hinw%(Index) = 1

End Sub

Private Sub hinw_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "hinw_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_hinw%(n%) = 0 Then Exit Sub
fl_hinw%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update usr_" & utabn(atyp(n%).Caption) & " set Hinweise='" + hinw(n%).text + "' where id='" + a_id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + a_id$ + "' and feldname='Hinweise'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
  cmd$ = "update auftritthigru set felddaten='" & hinw(n%).text & "' where auftrittsid='" & a_id$ & "' and feldname='Hinweise' and auftrittstyp='" & atyp(n%).Caption & "'"
Else
  neuid$ = form1.newid("auftritthigru", "id", 20)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten)" + _
         " values('" + neuid$ + "','" + _
           a_id$ + "','" + atyp(n%).Caption + "','Hinweise','" + hinw(n%).text + "')"
End If

Call form1.sqlqry(cmd$)

End Sub

Private Sub hlle_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "hlle_Change"
  If nochg% = 1 Then Exit Sub
  fl_hlle%(Index) = 1
End Sub

Private Sub hlle_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "hlle_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_hlle%(n%) = 0 Then Exit Sub
fl_hlle%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update usr_" & utabn(atyp(n%).Caption) & " set Halle='" + hlle(n%).text + "' where id='" + a_id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + a_id$ + "' and feldname='Halle'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
  cmd$ = "update auftritthigru set felddaten='" & hlle(n%).text & "' where auftrittsid='" & a_id$ & "' and feldname='Halle' and auftrittstyp='" & atyp(n%).Caption & "'"
Else
  neuid$ = form1.newid("auftritthigru", "id", 20)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten)" + _
         " values('" + neuid$ + "','" + _
           a_id$ + "','" + atyp(n%).Caption + "','Halle','" + hlle(n%).text + "')"
End If

Call form1.sqlqry(cmd$)


End Sub

Private Sub hon_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "hon_Change"
  If nochg% = 1 Then Exit Sub
  fl_hon%(Index) = 1
End Sub

Private Sub hon_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "hon_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_hon%(n%) = 0 Then Exit Sub
fl_hon%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update usr_" & LCase(atyp(n%).Caption) & " set Honorar='" + hon(n%).text + "' where id='" + a_id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + a_id$ + "' and feldname='Honorar'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
  cmd$ = "update auftritthigru set felddaten='" & hon(n%).text & "' where auftrittsid='" & a_id$ & "' and feldname='Honorar' and auftrittstyp='" & atyp(n%).Caption & "'"
Else
  neuid$ = form1.newid("auftritthigru", "id", 20)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten)" + _
         " values('" + neuid$ + "','" + _
           a_id$ + "','" + atyp(n%).Caption + "','Honorar','" + hon(n%).text + "')"
End If

Call form1.sqlqry(cmd$)

End Sub

Private Sub hsw_Click(Index As Integer)
Dim i%, sid$

'd2infile = "tpzoom": d2insub = "hsw_Click"
i% = Index
  sid$ = hlle(i%).text
  If Len(sid$) > 0 Then
    Load shwAdrDetail
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(sid$, "")
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
  Else
    Call lhlle_DblClick(i%)
  End If

End Sub

Private Sub lhlle_DblClick(Index As Integer)

'd2infile = "tpzoom": d2insub = "lhlle_DblClick"
  Load adrselect
  Call adrselect.sel_init("Halle", s$)
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then

  hlle(Index).text = adrselect.sel_getselected()
  Call hlle_LostFocus(Index)
  If trm(aort(Index).text) = "" Then
    o$ = ohnePLZ(form1.ortausadr("" & hlle(Index).text & ""))
    If o$ <> "" Then
      aort(Index).text = o$
      Call aort_LostFocus(Index)
    End If
  End If

  End If
  Unload adrselect

End Sub


Private Sub lveran_DblClick(Index As Integer)
'd2infile = "tpzoom": d2insub = "lveran_DblClick"
  Load adrselect
  Call adrselect.sel_init("Veranstalter", s$)
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then

  veran(Index).text = adrselect.sel_getselected()
  Call veran_LostFocus(Index)
  Unload adrselect

  End If

End Sub

Private Sub prg_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "prg_Change"
  If nochg% = 1 Then Exit Sub
  fl_prg%(Index) = 1
End Sub

Private Sub prg_Click(Index As Integer)
'd2infile = "tpzoom": d2insub = "prg_Click"
DoEvents
fl_prg%(Index) = 1
Call prg_LostFocus(Index)
End Sub

Private Sub prg_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "prg_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_prg%(n%) = 0 Then Exit Sub
fl_prg%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update usr_" & utabn(atyp(n%).Caption) + " set Programm='" + prg(n%).text + "' where id='" + a_id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + a_id$ + "' and feldname='Programm'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
  cmd$ = "update auftritthigru set felddaten='" & prg(n%).text & "' where auftrittsid='" & a_id$ & "' and feldname='Programm' and auftrittstyp='" & atyp(n%).Caption & "'"
Else
  neuid$ = form1.newid("auftritthigru", "id", 20)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten)" + _
         " values('" + neuid$ + "','" + _
           a_id$ + "','" + atyp(n%).Caption + "','Programm','" + prg(n%).text + "')"
End If

Call form1.sqlqry(cmd$)


End Sub

Private Sub shauf_Click(Index As Integer)
Dim id$
'd2infile = "tpzoom": d2insub = "shauf_Click"
id$ = aid(Index).Caption

Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(id$, 0)

End Sub

Private Sub styp_Click(Index As Integer)
'd2infile = "tpzoom": d2insub = "styp_Click"
atyp(Index).Enabled = True
atyp(Index).Visible = True
atyp(Index).Caption = styp(Index).List(styp(Index).ListIndex)
styp(Index).Enabled = False
styp(Index).Visible = False
Call form1.sqlqry("update auftritt set Auftrittstyp='" + atyp(Index).Caption + "' where id='" + aid(Index).Caption + "'")
id0$ = ""
Call Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Dim cp$, id$, idx%, r As ADODB.Recordset, s As ADODB.Recordset, n%, lck%
Dim rtmp As ADODB.Recordset, licount%, tid$

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "Timer1_Timer"
If tpzmode% = 0 Or tmrbusy% = 1 Then Exit Sub
'mode=1: read list from tplan.list6
'mode=2: read list from shwadrdetail.list2

nochg% = 1
If tpzmode% = 1 Then
  tpid$ = tplan.Text1(0).text
Else
  tpid$ = shwAdrDetail.datf(0).text
End If

cp$ = "Projektdetails " & tpid$
If Me.Caption <> cp$ Then
  Me.Caption = cp$
  On Error Resume Next
  Me.SetFocus
  On Error GoTo 0
End If

If tpzmode% = 1 Then
  licount% = tplan.List6.ListCount
Else
  licount% = shwAdrDetail.List2.ListCount
End If
If licount% = 0 Then
  nochg% = 0
  Exit Sub
End If

tmrbusy% = 1
If tpzmode% = 1 Then
  idx0% = tplan.List6.ListIndex
Else
  idx0% = shwAdrDetail.List2.ListIndex
End If
If idx0% < 0 Then idx0% = 0

If tpzmode% = 1 Then
  id$ = tplan.List6.List(idx0%)
Else
  id$ = shwAdrDetail.List2.List(idx0%)
End If
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
If id$ = id0$ Then
  tmrbusy% = 0
  nochg% = 0
  Exit Sub
End If
Call clrall
n% = 0
For idx% = idx0% To licount%
  If tpzmode% = 1 Then
    id$ = tplan.List6.List(idx%)
  Else
    id$ = shwAdrDetail.List2.List(idx%)
  End If
  id$ = Mid$(id$, InStr(id$, "(AID:") + 5)

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open "SELECT * FROM auftritt where id ='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly
    lck% = 1
    While Not r.EOF And n% < 5
      If r!id = id$ Then lck% = 0
      If lck% = 0 Then
        Select Case LCase$(r!auftrittstyp)
          Case "orchesterauftritt":
            hon(n%).Visible = True
            hlle(n%).Visible = True
            hsw(n%).Visible = True
            vsw(n%).Visible = True
            prg(n%).Visible = True
            prg(n%).Clear
            If tpzmode% = 1 Then
              For i% = 0 To tplan.List3.ListCount - 1
                tpl$ = tplan.List3.List(i%)
                tpl$ = trm(Left$(tpl$, InStr(tpl$, "    ")))
                prg(n%).AddItem tpl$
              Next i%
            End If
            hinw(n%).Visible = True
            veran(n%).Visible = True
          Case "künstlerauftritt":
            hon(n%).Visible = True
            hlle(n%).Visible = True
            hsw(n%).Visible = True
            vsw(n%).Visible = True
            prg(n%).Clear
            If tpzmode% = 1 Then
              For i% = 0 To tplan.List3.ListCount - 1
                tpl$ = tplan.List3.List(i%)
                tpl$ = trm(Left$(tpl$, InStr(tpl$, "    ")))
                prg(n%).AddItem tpl$
              Next i%
            End If
            prg(n%).Visible = True
            hinw(n%).Visible = True
            veran(n%).Visible = True
          Case Else:
        End Select
        adatum(n%).Caption = form1.dayofweek(r!datum) + ", " & datfromsql(r!datum) & " " & Left(r!zeit, 5) & " h"
        beg(n%).text = Left(trm(r!zeit), 5)
        atyp(n%).Caption = trm(r!auftrittstyp)
        If atyp(n%).Caption = transe("Neuer Auftritt") Then
          styp(n%).Clear
          styp(n%).Visible = True
          styp(n%).Enabled = True
          atyp(n%).Visible = False
          atyp(n%).Enabled = False
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT id FROM auftrittstypen order by sortierung", form1.adoc, dbOpenDynaset, dbReadOnly
          While Not rtmp.EOF
            styp(n%).AddItem rtmp!id
            rtmp.MoveNext
          Wend
          styp(n%).text = transe("Neuer Auftritt")
        End If
        If Not IsNull(r!bezeichnung) Then abez(n%).text = r!bezeichnung
        If Not IsNull(r!ort) Then aort(n%).text = r!ort
        aid(n%).Caption = r!id
        cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + r!id + "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
        While Not s.EOF
          Select Case LCase$(s!feldname)
            Case "halle"
              If Not IsNull(s!felddaten) Then hlle(n%).text = s!felddaten
            Case "veranstalter"
              If Not IsNull(s!felddaten) Then veran(n%).text = s!felddaten
            Case "programm"
              If Not IsNull(s!felddaten) Then prg(n%).text = s!felddaten
            Case "honorar"
              If Not IsNull(s!felddaten) Then hon(n%).text = s!felddaten
            Case "hinweise"
            If Not IsNull(s!felddaten) Then hinw(n%).text = s!felddaten
            Case Default:
          End Select
          s.MoveNext
        Wend
        n% = n% + 1
        's.Close
      End If
      r.MoveNext
    Wend
    If idx% = idx0% Then
      id0$ = id$
    End If
    If n% > 5 Then Exit For
Next idx%
On Error Resume Next
Me.SetFocus
On Error GoTo 0
nochg% = 0
tmrbusy% = 0
End Sub

Sub clrall()
Dim n%

'd2infile = "tpzoom": d2insub = "clrall"
For n% = 0 To 4
  atyp(n%).Visible = True
  atyp(n%).Enabled = True
  styp(n%).Visible = False
  styp(n%).Enabled = False
  fl_abez%(n%) = 0
  fl_ort%(n%) = 0
  fl_hon%(n%) = 0
  fl_hlle%(n%) = 0
  fl_prg%(n%) = 0
  fl_hinw%(n%) = 0
  fl_veran%(n%) = 0
  adatum(n%).Caption = ""
  atyp(n%).Caption = ""
  aort(n%).text = ""
  abez(n%).text = ""
  hon(n%).text = ""
  aid(n%).Caption = ""
  hlle(n%).text = ""
  prg(n%).text = ""
  hinw(n%).text = ""
  veran(n%).text = ""
hon(n%).Visible = False
hlle(n%).Visible = False
hsw(n%).Visible = False
vsw(n%).Visible = False
prg(n%).Visible = False
hinw(n%).Visible = False
veran(n%).Visible = False

Next n%
End Sub

Private Sub veran_Change(Index As Integer)
'd2infile = "tpzoom": d2insub = "veran_Change"
  If nochg% = 1 Then Exit Sub
  fl_veran%(Index) = 1
End Sub

Private Sub veran_LostFocus(Index As Integer)
Dim a_id$, n%, cmd$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tpzoom": d2insub = "veran_LostFocus"
If nochg% = 1 Then Exit Sub
n% = Index
If fl_veran%(n%) = 0 Then Exit Sub
fl_veran%(n%) = 0
a_id$ = aid(n%).Caption
cmd$ = "update usr_" & utabn(atyp(n%).Caption) + " set Veranstalter='" + veran(n%).text + "' where id='" + a_id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "SELECT * FROM auftritthigru where auftrittsid ='" + a_id$ + "' and feldname='Veranstalter'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
s.Open cmd$, form1.adoc, dbOpenDynaset, dbReadOnly
If Not s.EOF Then
  cmd$ = "update auftritthigru set felddaten='" & veran(n%).text & "' where auftrittsid='" & a_id$ & "' and feldname='Veranstalter' and auftrittstyp='" & atyp(n%).Caption & "'"
Else
  neuid$ = form1.newid("auftritthigru", "id", 20)
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten)" + _
         " values('" + neuid$ + "','" + _
           a_id$ + "','" + atyp(n%).Caption + "','Veranstalter','" + veran(n%).text + "')"
End If

Call form1.sqlqry(cmd$)

End Sub

Private Sub vsw_Click(Index As Integer)
Dim i%, sid$

'd2infile = "tpzoom": d2insub = "vsw_Click"
i% = Index
  sid$ = veran(i%).text
  If Len(sid$) > 0 Then
    Load shwAdrDetail
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(sid$, "")
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
  Else
    Call lveran_DblClick(i%)
  End If


End Sub
Public Sub setmode(M%)
'd2infile = "tpzoom": d2insub = "setmode"
tpzmode% = M%
Call Timer1_Timer
End Sub
