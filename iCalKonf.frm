VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form iCalKonf 
   Caption         =   "iKalender Konfiguration"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton delterm 
      Caption         =   "Alle Termine löschen"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Alles aktualisieren"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Benutzer anlegen"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Termintypen übertragen"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      Picture         =   "iCalKonf.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "iCalKonf.frx":03A7
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   1560
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   5895
   End
End
Attribute VB_Name = "iCalKonf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim rtmp As ADODB.Recordset, c$
Dim rrr, i%, dbn$, co$

dbn$ = form1.getdbname()
c$ = "delete from webcal_categories"
Call form1.kalsqlqry(c$)

c$ = "select MAX(cat_id) as maxi from webcal_categories"
i% = 1
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c$, form1.kaldb, adOpenDynamic, adLockReadOnly
If rtmp.EOF Then
  Command2.Caption = "Fehler: id nicht zu ermitteln"
  Exit Sub
End If
If Not IsNull(rtmp!maxi) Then
  i% = rtmp!maxi + 1
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftrittstypen order by sortierung", form1.adoc, dbOpenDynaset, dbReadOnly)
While Not rtmp.EOF
  Command2.Caption = rtmp!id: DoEvents
  On Error Resume Next
  co$ = "#" + hex2(rtmp!Kalenderfarbe_R) + hex2(rtmp!Kalenderfarbe_G) + hex2(rtmp!Kalenderfarbe_B)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then co$ = "#000000"
  c$ = "insert into webcal_categories (cat_id,cat_name,cat_color) values("
  c$ = c$ + trm(i%) + ",'" + trm(rtmp!id) + "','" + co$ + "')"
'Debug.Print c$
  Call form1.kalsqlqry(c$)
  i% = i% + 1
  rtmp.MoveNext
Wend
Command2.Caption = "Termintypen übertragen"

End Sub

Private Sub Command3_Click()
Dim rtmp As ADODB.Recordset, c$
Dim s As ADODB.Recordset
Dim rrr, dbn$, co$, cn$, kid$, vid$

c$ = "delete from webcal_user where cal_is_admin='N'"
Call form1.kalsqlqry(c$)

dbn$ = form1.getdbname()
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT vid, kid from adresstyp WHERE typ='webcal'", form1.adoc, dbOpenDynaset, dbReadOnly)
While Not rtmp.EOF
  kid$ = trm(rtmp!kid)
  vid$ = trm(rtmp!vid)
  If kid$ <> "-1" And kid$ <> "" Then
    cn$ = form1.getkontaktnamebyid(kid$)
  Else
    cn$ = form1.getnamebyid(vid$)
  End If
  Command3.Caption = cn$: DoEvents
  c$ = "insert into webcal_user (cal_login,cal_passwd,cal_lastname,cal_firstname,cal_is_admin,cal_email,cal_enabled) values("
  co$ = form1.higruget(trm(rtmp!vid), trm(rtmp!kid), "webcal", "Benutzername")
  c$ = c$ + "'" + co$ + "'"
  co$ = form1.higruget(trm(rtmp!vid), trm(rtmp!kid), "webcal", "Passwort")
  c$ = c$ + ",'" + co$ + "'"
  c$ = c$ + ",'" + word1(cn$) + "'"
  c$ = c$ + ",'" + word2bis(cn$) + "','N'"
  If kid$ <> "-1" And kid$ <> "" Then
    co$ = form1.getkontaktemailbyid(kid$)
  Else
    co$ = form1.getemailbyid(vid$)
  End If
  c$ = c$ + ",'" + word1(co$) + "','Y')"
Debug.Print c$
  Call form1.kalsqlqry(c$)
  rtmp.MoveNext
Wend
Command3.Caption = "Benutzer anlegen"


End Sub

Private Sub Command4_Click()
Call Command2_Click
Call Command3_Click
End Sub

Private Sub delterm_Click()
Dim c$

   c$ = "delete from webcal_entry": Call form1.kalsqlqry(c$)
   c$ = "delete from webcal_entry_user": Call form1.kalsqlqry(c$)
   c$ = "delete from webcal_entry_log": Call form1.kalsqlqry(c$)
   c$ = "delete from webcal_entry_categories": Call form1.kalsqlqry(c$)

End Sub

Private Sub Form_Load()
Dim cf$
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub
Private Sub Form_Resize()
axsResizer1.Resize
End Sub

