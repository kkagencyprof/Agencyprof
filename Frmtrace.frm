VERSION 5.00
Begin VB.Form frmTrace 
   Caption         =   "Tracing Options"
   ClientHeight    =   2265
   ClientLeft      =   6255
   ClientTop       =   5370
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4920
   Begin VB.Frame Frame1 
      Caption         =   "Tracing Level"
      Height          =   855
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Width           =   4575
      Begin VB.OptionButton OptLevel 
         Caption         =   "4:  Verbose Trace"
         Height          =   255
         Index           =   4
         Left            =   2460
         TabIndex        =   9
         Top             =   540
         Width           =   1875
      End
      Begin VB.OptionButton OptLevel 
         Caption         =   "2:  Warning Winsock Calls"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Width           =   2235
      End
      Begin VB.OptionButton OptLevel 
         Caption         =   "1:  Error Winsock Calls"
         Height          =   255
         Index           =   1
         Left            =   2460
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptLevel 
         Caption         =   "0:  All Winsock Calls"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.TextBox txtTraceFile 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.CheckBox chkEnableTracing 
      Caption         =   "Enable Tracing"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Trace File:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   795
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bApply As Boolean 'Determines if the user clicked cancel or OK
Private bEnableTracingSave As Boolean
Private nFlagSave As Integer
Private strTraceFileSave As String


Private Sub cmdCancel_Click()
    bApply = False
    If bEnableTracingSave Then
        chkEnableTracing.value = 1
    Else
        chkEnableTracing.value = 0
    End If

    OptLevel(nFlagSave).value = True

    txtTraceFile.Text = strTraceFileSave

    Me.Hide
End Sub

Private Sub cmdOK_Click()
    bApply = True
    Me.Hide
End Sub

Public Sub TraceParams(bEnableTracing As Boolean, nFlag As Integer, strTraceFile As String)

    Dim nIndex As Integer

    If (chkEnableTracing.value = 1) Then
        bEnableTracing = True
    Else
        bEnableTracing = False
    End If

    For nIndex = 0 To 4
        If nIndex <> 3 Then
            If (OptLevel(nIndex).value) Then nFlag = nIndex
        End If
    Next nIndex

    strTraceFile = txtTraceFile.Text

End Sub

Private Sub Form_Activate()

    Dim nIndex As Integer

    If (chkEnableTracing.value = 1) Then
        bEnableTracingSave = True
    Else
        bEnableTracingSave = False
    End If

    For nIndex = 0 To 4
        If nIndex <> 3 Then
            If (OptLevel(nIndex).value) Then nFlagSave = nIndex
        End If
    Next nIndex

    strTraceFileSave = txtTraceFile.Text

End Sub
