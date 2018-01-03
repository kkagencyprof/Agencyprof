VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form frmBrowser 
   ClientHeight    =   10560
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox locdialer 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5520
      Top             =   1500
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Oben ausrichten
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11175
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   11175
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   300
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Adresse:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Adresse:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Oben ausrichten
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1376
      ButtonWidth     =   1720
      ButtonHeight    =   1217
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zurück"
            Key             =   "Back"
            Object.ToolTipText     =   "Zurück zur vorherigen Seite"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vor"
            Key             =   "Forward"
            Object.ToolTipText     =   "Weiter zur nächsten Seite"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abbrechen"
            Key             =   "Stop"
            Object.ToolTipText     =   "Anhalten"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Neu laden"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Aktualisieren"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anfang"
            Key             =   "Home"
            Object.ToolTipText     =   "Startseite"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Suchen"
            Key             =   "Search"
            Object.ToolTipText     =   "Durchsuchen"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tschüs"
            Key             =   "culater"
            Object.ToolTipText     =   "Browser schliessen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2A6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5400
      ExtentX         =   9513
      ExtentY         =   6586
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Form_Load()
d2infile = "frmBrowser": d2insub = "Form_Load"
    On Error Resume Next
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Width = form1.mylastwidth(Me.name, 1)
Me.Height = form1.mylastheight(Me.name, 1)
Call form1.formpos(Me)

lblAddress.Caption = transe("&Adresse:")
tbToolBar.Buttons(1).Caption = transe("Zurück")
tbToolBar.Buttons(1).ToolTipText = transe("Zurück zur vorherigen Seite")
tbToolBar.Buttons(2).Caption = transe("Voraus")
tbToolBar.Buttons(2).ToolTipText = transe("Weiter zur nächsten Seite")
tbToolBar.Buttons(3).Caption = transe("Abbrechen")
tbToolBar.Buttons(3).ToolTipText = transe("Anhalten")
tbToolBar.Buttons(4).Caption = transe("Neu laden")
tbToolBar.Buttons(4).ToolTipText = transe("Aktualisieren")
tbToolBar.Buttons(5).Caption = transe("Anfang")
tbToolBar.Buttons(5).ToolTipText = transe("Startseite")
tbToolBar.Buttons(6).Caption = transe("Suchen")
tbToolBar.Buttons(6).ToolTipText = transe("Durchsuchen")
tbToolBar.Buttons(7).Caption = transe("Tschüs")
tbToolBar.Buttons(7).ToolTipText = transe("Browser schliessen")

If Not form1.brwhidden Then Me.Show
    
    tbToolBar.Refresh
    Form_Resize


    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15


    If Len(StartingAddress) > 0 Then Call setstarturl(StartingAddress)

End Sub

Public Sub setstarturl(url As String)
        cboAddress.text = url
        cboAddress.AddItem cboAddress.text
        'versuche auf Startadresse zu positionieren
        timTimer.Enabled = True
        brwWebBrowser.Navigate url
End Sub

Private Sub brwWebBrowser_DownloadComplete()
Dim t$

    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
'    Debug.Print brwWebBrowser.Document.Body.innerhtml
End Sub


Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, url As Variant)
d2infile = "frmBrowser": d2insub = "brwWebBrowser_NavigateComplete2"
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Private Sub cboAddress_Click()
d2infile = "frmBrowser": d2insub = "cboAddress_Click"
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
d2infile = "frmBrowser": d2insub = "cboAddress_KeyPress"
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub


Private Sub Form_Resize()
d2infile = "frmBrowser": d2insub = "Form_Resize"
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 200
    brwWebBrowser.Left = 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub


Private Sub Form_Unload(Cancel As Integer)
d2infile = "frmBrowser": d2insub = "Form_Unload"
If trm(locdialer.text) <> "" Then
  On Error Resume Next
  Kill locdialer.text
  On Error GoTo 0
End If
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)
exuld:
On Error GoTo 0
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim zu$

d2infile = "frmBrowser": d2insub = "tbToolBar_ButtonClick"
    On Error Resume Next


    timTimer.Enabled = True


    Select Case Button.key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
                timTimer.Enabled = True
                zu$ = form1.getmyhomepg()
                brwWebBrowser.Navigate zu$
        Case "Search"
                timTimer.Enabled = True
                brwWebBrowser.Navigate "http://www.google.de"
        Case "Stop"
            If (CtrlKey()) Then
              Load MultiPList
              MultiPList.Text1.text = brwWebBrowser.Document.Body.innerhtml
            Else
              timTimer.Enabled = False
              brwWebBrowser.Stop
              Me.Caption = brwWebBrowser.LocationName
            End If
        Case "culater"
            Unload Me
    End Select

End Sub

Private Sub timTimer_Timer()
d2infile = "frmBrowser": d2insub = "timTimer_Timer"
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Bearbeiten..."
    End If
End Sub


