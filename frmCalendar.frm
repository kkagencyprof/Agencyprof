VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmCalendar.frx":0000
      Left            =   720
      List            =   "frmCalendar.frx":0002
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "2000"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   360
      Picture         =   "frmCalendar.frx":0004
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgClose 
      Height          =   240
      Left            =   120
      Picture         =   "frmCalendar.frx":014E
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label labToday 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   45
   End
   Begin VB.Image imgArrRight 
      Height          =   240
      Left            =   360
      Picture         =   "frmCalendar.frx":0298
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgArrLeft 
      Height          =   240
      Left            =   120
      Picture         =   "frmCalendar.frx":03E2
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cal_dtrenn As String
'==== FlatButtons ====

Private Declare Function DrawIconEx _
  Lib "user32" _
  (ByVal hdc As Long, _
   ByVal xLeft As Long, ByVal yTop As Long, _
   ByVal hicon As Long, _
   ByVal cxWidth As Long, ByVal cyWidth As Long, _
   ByVal istepIfAniCur As Long, _
   ByVal hbrFlickerFreeDraw As Long, _
   ByVal diFlags As Long) As Long
Private Const DI_MASK As Long = 1&
Private Const DI_IMAGE As Long = 2&
Private Const DI_NORMAL As Long = DI_MASK Or DI_IMAGE

Private Enum FB_STATE
  FB_FLAT
  FB_NORMAL
  FB_PRESSED
  FB_DISABLED
End Enum

'==== Allgemein ====

Private Type Area
  X As Long
  Y As Long
  b As Long
  h As Long
End Type
Private MyArea(7) As Area
Private DayArea As Area

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function ClientToScreen _
  Lib "user32" _
  (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long

Private Const A_MONTH_L As Long = 0
Private Const A_MONTH_R As Long = 1
Private Const A_YEAR_L As Long = 2
Private Const A_YEAR_R As Long = 3
Private Const A_WEEK_L As Long = 4
Private Const A_WEEK_R As Long = 5
Private Const A_EXIT As Long = 6
Private Const A_OK As Long = 7

Private WeekDayStr(6) As String
Private Const fButB As Long = 11
Private lrB As Long 'Breite eines LiRe-Controls
Private LineH As Long
Private LineY As Long
Private DownBut As Long
Private DayDownX As Long
Private DayDownY As Long

Private TWIX As Long, TWIY As Long
Private aDate As Date 'aktuelles Datum
Private tDate As Date 'Tabellen-Startdatum
Private CMB_BLOCKED As Boolean
Private TXT_BLOCKED As Boolean
Private SelOK As Boolean
Private MouseDownX As Long, MouseDownY As Long


Private Sub Form_Load()
  SelOK = False
  DownBut = -1
End Sub

'==== PUBLIC ====================

Public Property Get SelectionOK() As Boolean
  SelectionOK = SelOK
End Property

Public Property Get SelectedDate() As Date
  SelectedDate = aDate
End Property

Public Sub Init(ByVal obj As Object, _
                ByVal sDate As String, _
                Optional ByVal xDiff As Long, _
                Optional ByVal yDiff As Long)
  Dim p As POINTAPI, frmt$
  Dim i As Long

  cal_dtrenn = form1.dttrenn
  '
  '---- Konstanten initialisieren
  'Twips <-> Pixel
  TWIX = Screen.TwipsPerPixelX
  TWIY = Screen.TwipsPerPixelY
  'Wochentage Kurzform (Mo - So)
  For i = 0 To 6
    WeekDayStr(i) = Format(CDate((3 + i) & cal_dtrenn & "1" & cal_dtrenn & "2000"), "ddd")
  Next
  '==== Fenster Position festlegen
  With obj
    p.X = ScaleX(.Left, .Parent.ScaleMode, vbPixels)
    p.Y = ScaleY(.Top + .Height, .Parent.ScaleMode, vbPixels)
    ClientToScreen .Parent.hWnd, p
    p.X = p.X * TWIX + ScaleX(xDiff, .Parent.ScaleMode, vbTwips)
    p.Y = p.Y * TWIY + ScaleX(yDiff, .Parent.ScaleMode, vbTwips)
  End With
  
  '==== Eingangsdatum
  If Not IsDate(sDate) Then sDate = word1(Now())
  aDate = CDate(sDate)
  If (apyear(aDate) = 9999) Then aDate = CDate("31" & cal_dtrenn & "12" & cal_dtrenn & "9998")
  If (apyear(aDate) = 100) Then aDate = CDate("1" & cal_dtrenn & "1" & cal_dtrenn & "101")
'Call form1.dbg2f("frmCalendar.Init(...): adate=" + trm(aDate))
  
  '==== Controls aufbauen
  '---- Voreinstellungen
  TXT_BLOCKED = True
  txtYear = apyear(aDate)
  TXT_BLOCKED = False
  lrB = fButB * 2 + 2 'Breite eines LiRe-Controls
  With cmbMonth
    '---- Zeilenhöhe
    LineH = .Height 'Netto Zeilenhöhe
    LineY = (LineH - TextHeight("X")) \ 2 'Text-Y-Offset
    '---- Fenstergröße
    i = Width \ TWIX - ScaleWidth 'Rahmen links/rechts
    Width = (lrB + 7 * LineH + 7 + i) * TWIX
    i = Height \ TWIY - ScaleHeight 'Rahmen oben/unten
    Height = (9 * LineH + 9 + labToday.Height + i) * TWIY
    '---- checken, ob unten genug Platz ist
    If ((p.Y + Height) > Screen.Height) Then
      p.Y = p.Y - Height - obj.Height
    End If
    '---- checken, ob rechts genug Platz ist
    If ((p.X + Width) > Screen.Width) Then p.X = Screen.Width - Width
    '---- checken, ob links genug Platz ist
    If (p.X < 0) Then p.X = 0
    '---- Position setzen
    Left = p.X
    Top = p.Y
    '---- Controlhöhe
    txtYear.Height = LineH
    '---- Monatsnamen
    For i = 1 To 12
      .AddItem Format(CDate("1" & cal_dtrenn & i & cal_dtrenn & "2000"), "mmmm")
    Next
    '---- aktuellen Monat voreinstellen
    CMB_BLOCKED = True
    .ListIndex = apmonth(aDate) - 1
    CMB_BLOCKED = False
    '---- Zeile 'Today' positionieren
    labToday.Top = ScaleHeight - labToday.Height
  End With
  '---- Zeile 'Today'
  frmt$ = "dddd, dd. mmmm yyyy"
  With labToday
    .Caption = Format(Date, frmt$)
    .Left = (ScaleWidth - .Width) \ 2 'zentrieren
  End With
  '---- 1.Zeile
  'Month re/li
  DrawAreaEx 0, 0, lrB, LineH, FB_PRESSED
  'FormControls
  i = ScaleWidth - LineH - 1
  'Year re/li
  i = i - lrB
  DrawAreaEx i, 0, lrB, LineH, FB_PRESSED
  'Text Year
  txtYear.Move i - txtYear.Width - 1, 0
  'Combo Month
  cmbMonth.Move lrB + 1, 0
  cmbMonth.Width = txtYear.Left - cmbMonth.Left - 1
  '---- 2.Zeile
  'Week re/li
  DrawAreaEx 0, LineH + 1, lrB, LineH, FB_PRESSED
  'Rahmen Wochennummern
  DrawAreaEx lrB + 1, LineH + 1, 7 * LineH + 6, LineH, FB_PRESSED
  
  '---- ButtonAreas (LinksRechts)
  With MyArea(A_MONTH_L)
    .X = 1: .Y = 1: .b = fButB: .h = LineH - 2
  End With
  With MyArea(A_MONTH_R)
    .X = 1 + fButB: .Y = 1: .b = fButB: .h = LineH - 2
  End With
  With MyArea(A_YEAR_L)
    .X = i + 1: .Y = 1: .b = fButB: .h = LineH - 2
  End With
  With MyArea(A_YEAR_R)
    .X = i + 1 + fButB: .Y = 1: .b = fButB: .h = LineH - 2
  End With
  With MyArea(A_WEEK_L)
    .X = 1: .Y = LineH + 2: .b = fButB: .h = LineH - 2
  End With
  With MyArea(A_WEEK_R)
    .X = 1 + fButB: .Y = LineH + 2: .b = fButB: .h = LineH - 2
  End With
  '---- FormControlAreas
  With MyArea(A_EXIT)
    .X = i + lrB + 1: .Y = 0: .b = LineH: .h = LineH \ 2
  End With
  With MyArea(A_OK)
    .X = i + lrB + 1: .Y = LineH \ 2: .b = LineH: .h = LineH \ 2
  End With
  '---- FlatButtons setzen
  For i = 0 To 7
    DrawArea MyArea(i), FB_NORMAL
    If (i <= 5) Then
      SetIcon IIf(i And 1, imgArrRight, imgArrLeft), MyArea(i), 16
    Else
      SetIcon IIf((i = 6), imgClose, imgOK), MyArea(i), 16
    End If
  Next
  '---- TagNamen (Mo-So)
  For i = 0 To 6
    CurrentX = (lrB - TextWidth(WeekDayStr(i))) \ 2 ' + 1
    CurrentY = LineH * 2 + 2 + i * (LineH + 1) + LineY
    If (i > 4) Then ForeColor = vbRed
    Print WeekDayStr(i)
  Next
  '---- Rahmen TagNamen
  DrawAreaEx 0, 2 * LineH + 2, lrB, 7 * LineH + 6, FB_PRESSED
  '---- Area Tage
  With DayArea
    .X = lrB + 1
    .Y = 2 * LineH + 2
    .b = 7 * LineH + 6
    .h = .b
  End With
  '---- Anzeigen
  DrawDays
End Sub

'==== CONTROLS ====================

Private Sub cmbMonth_Click()
Dim rrr, aDtt As String

  If (CMB_BLOCKED) Then Exit Sub
  On Error Resume Next
  aDate = CDate(apday(aDate) & cal_dtrenn & (cmbMonth.ListIndex + 1) & cal_dtrenn & apyear(aDate))
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    aDtt = (apday(aDate) - 1) & cal_dtrenn & (cmbMonth.ListIndex + 1) & cal_dtrenn & apyear(aDate)
    On Error Resume Next
    aDate = CDate(aDtt)
    On Error GoTo 0
  End If
  TXT_BLOCKED = True
  txtYear = trm(str(apyear(aDate)))
  TXT_BLOCKED = False
  DrawDays
End Sub

Private Sub txtYear_Change()
  Dim i As Long
  '
  If (TXT_BLOCKED) Then Exit Sub
  i = Val(txtYear)
  If (i < 101) Then Exit Sub
  If (i > 9998) Then Exit Sub
  aDate = CDate(apday(aDate) & cal_dtrenn & apmonth(aDate) & cal_dtrenn & i)
  DrawDays
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
  If (KeyAscii > 31) Then
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
  End If
End Sub

Private Sub labToday_Click()
  If (aDate <> Date) Then
    aDate = Date
    RefreshDate
  End If
End Sub

Private Sub labToday_DblClick()
  aDate = Date
  SelOK = True: Hide
End Sub

'==== STUFF ====================

Private Sub RefreshDate()
  CMB_BLOCKED = True
  cmbMonth.ListIndex = apmonth(aDate) - 1
  CMB_BLOCKED = False
  TXT_BLOCKED = True
  txtYear = trm(str(apyear(aDate)))
  TXT_BLOCKED = False
  DrawDays
End Sub

Private Sub DrawDays()
  Dim Dt As Date
  Dim s As String
  Dim w As Long, t As Long
  '
  '---- Datum des Tabellenanfangs ermitteln
  Dt = DateAdd("d", _
               -((Weekday(aDate) + 5) Mod 7), aDate) 'auf Mo setzen
  Dt = DateAdd("ww", -3, Dt) 'auf Tabellenanfang setzen
  tDate = Dt 'und global merken
  '==== Platz schaffen
  '---- Wochennummern
  Me.Line (lrB + 2, LineH + 2)- _
          (ScaleWidth - 2, LineH + LineH - 1), _
          vbButtonFace, BF
  '---- Wochentage
  Me.Line (lrB + 2, 2 * LineH + 3)- _
          (ScaleWidth - 2, ScaleHeight - labToday.Height - 1), _
          vbButtonFace, BF
  
  '==== Anzeigen
  w = 0
  Do
    '---- Wochennummer
    ForeColor = vbBlack
    FontBold = False
    s = trm(str(WeekNumber(Dt)))
    CurrentX = lrB + 1 + _
               w * (LineH + 1) + (LineH - TextWidth(s)) \ 2
    CurrentY = LineH + 1 + LineY
    Print s
    '---- Wochentage
    t = 0
    Do
      If (Dt = aDate) Then 'aktueller Tag
        ForeColor = vbWhite
      Else 'anderer Tag
        ForeColor = IIf(t > 4, vbRed, vbBlack)
      End If
      FontBold = (apmonth(Dt) = apmonth(aDate))
      s = Format(Dt, "d")
      CurrentX = lrB + 1 + _
                 w * (LineH + 1) + (LineH - TextWidth(s)) \ 2
      CurrentY = (t + 2) * (LineH + 1) + LineY
      Print s
      '---- nächster Tag
      t = t + 1: Dt = DateAdd("d", 1, Dt)
    Loop Until (t > 6)
    '---- nächste Woche
    w = w + 1
  Loop Until (w > 6)
  Refresh
End Sub

Private Function WeekNumber(ByVal d As Date) As Long
  Dim i As Long, rrr, wd
  '
  WeekNumber = Val(Format(d, "ww", vbMonday))
  '---- falls der 1.1. > Donnerstag ist -> noch letzte Dez-Woche
  If ((Weekday(CDate("1" & cal_dtrenn & "1" & cal_dtrenn & apyear(d)) + 5) Mod 7) > 3) Then
    WeekNumber = WeekNumber - 1 '(CDate("31.12." & apyear(d) - 1))
  End If
  If (apmonth(d) = 12) Then
    If (apday(d) > 28) Then
      '---- falls der 31.12. < Donnerstag ist
      '---- -> bereits 1. Jan-Woche
      On Error Resume Next
      wd = (Weekday(CDate("31" & cal_dtrenn & "12" & cal_dtrenn & apyear(d)) + 5) Mod 7)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then wd = (Weekday(CDate("31/12/" & apyear(d)) + 5) Mod 7)
      If (wd < 3) Then
        WeekNumber = 1
      End If
    End If
  End If
End Function

Private Function PointInArea(ByVal px As Long, _
                             ByVal py As Long, _
                             ByRef a As Area) As Boolean
'liefert TRUE wenn p in a ist, sonst FALSE
  PointInArea = False
  With a
    If (px < .X) Then Exit Function
    If (px >= (.X + .b)) Then Exit Function
    If (py < .Y) Then Exit Function
    If (py >= (.Y + .h)) Then Exit Function
  End With
  PointInArea = True
End Function

Private Function PointedArea(ByVal px As Long, _
                             ByVal py As Long) As Long
  Dim i As Long
  '
  For i = 0 To 7
    If (PointInArea(px, py, MyArea(i))) Then
      PointedArea = i: Exit Function
    End If
  Next
  PointedArea = -1
End Function

Private Sub DrawArea(ByRef a As Area, _
                     ByVal s As FB_STATE)
  Dim r As RECT
  '
  With r
    .Left = a.X: .Top = a.Y
    .Right = a.X + a.b: .Bottom = a.Y + a.h
  End With
  Select Case (s)
    Case FB_FLAT:
      DrawEdge hdc, r, BDR_RAISEDINNER, BF_RECT Or BF_FLAT
    Case FB_NORMAL:
      DrawEdge hdc, r, BDR_RAISEDINNER, BF_RECT
    Case FB_PRESSED:
      DrawEdge hdc, r, BDR_SUNKENOUTER, BF_RECT
    Case FB_DISABLED:
      DrawEdge hdc, r, BDR_RAISEDINNER, BF_RECT Or BF_SOFT
  End Select
End Sub

Private Sub DrawAreaEx(ByVal X As Long, _
                       ByVal Y As Long, _
                       ByVal b As Long, _
                       ByVal h As Long, _
                       ByVal s As FB_STATE)
  Dim a As Area
  '
  With a
    .X = X: .Y = Y: .b = b: .h = h
  End With
  DrawArea a, s
End Sub

Private Sub SetIcon(ByVal i As Image, _
                    ByRef a As Area, _
                    ByVal k As Long)
  With a
    DrawIconEx hdc, _
               .X + (.b - k) \ 2, _
               .Y + (.h - k) \ 2, _
               i.Picture.Handle, _
               0, 0, 0, 0, DI_NORMAL
  End With
End Sub

Private Sub ExecuteBut(ByVal Index As Long)
  Dim i As Long
  '
  Select Case (Index)
    Case A_MONTH_L:
      If (apmonth(aDate) = 1) And (apyear(aDate) = 101) Then Exit Sub
      aDate = DateAdd("m", -1, aDate)
      cmbMonth.ListIndex = apmonth(aDate) - 1
    Case A_MONTH_R:
      If (apmonth(aDate) = 12) And (apyear(aDate) = 9998) Then Exit Sub
      aDate = DateAdd("m", 1, aDate)
      cmbMonth.ListIndex = apmonth(aDate) - 1
    Case A_YEAR_L:
      If (apyear(aDate) = 101) Then Exit Sub
      aDate = DateAdd("yyyy", -1, aDate)
      TXT_BLOCKED = True
      txtYear = trm(str(apyear(aDate)))
      TXT_BLOCKED = False
      DrawDays
    Case A_YEAR_R:
      If (apyear(aDate) = 9998) Then Exit Sub
      aDate = DateAdd("yyyy", 1, aDate)
      TXT_BLOCKED = True
      txtYear = trm(str(apyear(aDate)))
      TXT_BLOCKED = False
      DrawDays
    Case A_WEEK_L:
      If (apmonth(aDate) = 1) And (apyear(aDate) = 101) Then Exit Sub
      aDate = DateAdd("ww", -1, aDate)
      RefreshDate
    Case A_WEEK_R:
      If (apmonth(aDate) = 12) And (apyear(aDate) = 9998) Then Exit Sub
      aDate = DateAdd("ww", 1, aDate)
      RefreshDate
    Case A_EXIT:
      Hide
    Case A_OK:
      SelOK = True: Hide
  End Select
End Sub

'==== FORM ====================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift And vbCtrlMask) > 0 Then
    Select Case (KeyCode)
      '---- Tag wechseln
      Case 37: 'Left
        If (apyear(aDate) > 100) Then
          aDate = DateAdd("ww", -1, aDate)
          RefreshDate
          KeyCode = 0
        End If
      Case 39: 'Right
        If (apyear(aDate) < 9999) Then
          aDate = DateAdd("ww", 1, aDate)
          RefreshDate
          KeyCode = 0
        End If
      Case 38: 'Up
        If (apyear(aDate) > 100) Then
          aDate = DateAdd("d", -1, aDate)
          RefreshDate
          KeyCode = 0
        End If
      Case 40: 'Down
        If (apyear(aDate) < 9999) Then
          aDate = DateAdd("d", 1, aDate)
          RefreshDate
          KeyCode = 0
        End If
      '---- Monat/Jahr wechseln
      Case 33: 'PgUp
        If (Shift And vbShiftMask) > 0 Then 'Jahr zurück
          If (apyear(aDate) > 101) Then
            aDate = DateAdd("yyyy", -1, aDate)
            RefreshDate
            KeyCode = 0
          End If
        Else 'Monat zurück
          If (apyear(aDate) > 100) Then
            aDate = DateAdd("m", -1, aDate)
            RefreshDate
            KeyCode = 0
          End If
        End If
      Case 34: 'PgDown
        If (Shift And vbShiftMask) > 0 Then 'Jahr weiter
          If (apyear(aDate) < 9998) Then
            aDate = DateAdd("yyyy", 1, aDate)
            RefreshDate
            KeyCode = 0
          End If
        Else 'Monat weiter
          If (apyear(aDate) < 9999) Then
            aDate = DateAdd("m", 1, aDate)
            RefreshDate
            KeyCode = 0
          End If
        End If
    End Select
  End If
  If (KeyCode = 72) Then labToday_Click 'h/H
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Select Case (KeyAscii)
    Case vbKeyReturn:
      SelOK = True: Hide
    Case vbKeyEscape:
      Hide
  End Select
End Sub

Private Sub Form_DblClick()
  Dim bu As Integer, sh As Integer, X As Single, Y As Single
  '
  If (PointInArea(MouseDownX, MouseDownY, DayArea)) Then
    SelOK = True: Hide
  Else
    '---- zweiten Click weiterleiten
    bu = 1: sh = 0: X = MouseDownX: Y = MouseDownY
    Form_MouseDown bu, sh, X, Y
    Form_MouseUp bu, sh, X, Y
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, Y As Single)
  Dim i As Long
  '
  If Not (Button = 1) Then Exit Sub
  MouseDownX = X: MouseDownY = Y
  '---- FlatButton-Areas testen
  i = PointedArea(X, Y)
  If (i > -1) Then
    DownBut = i
    DrawArea MyArea(i), FB_PRESSED: Refresh
  '---- Tage-Area testen
  ElseIf (PointInArea(X, Y, DayArea)) Then
    'in den Zwischenraum?
    If ((X - DayArea.X) Mod (LineH + 1)) = LineH Then Exit Sub
    If ((Y - DayArea.Y) Mod (LineH + 1)) = LineH Then Exit Sub
    'sonst DownDay merken
    DayDownX = (X - DayArea.X) \ (LineH + 1)
    DayDownY = (Y - DayArea.Y) \ (LineH + 1)
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, Y As Single)
  Dim i As Long
  '
  If (DownBut > -1) Then
    'Testen, ob Maus den gedrückten Button verlassen hat
    i = PointedArea(X, Y)
    If (i <> DownBut) Then
      DrawArea MyArea(DownBut), FB_NORMAL: Refresh
      DownBut = -1
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         X As Single, Y As Single)
  Dim i As Long
  '
  If (Button = 2) Then Hide
  '---- FlatButton-Areas testen
  i = PointedArea(X, Y)
  If (i > -1) Then
    If (i = DownBut) Then
      DrawArea MyArea(DownBut), FB_NORMAL: Refresh
      DownBut = -1
      ExecuteBut i
    End If
  '---- Tage-Area testen
  ElseIf (PointInArea(X, Y, DayArea)) Then
    'in den Zwischenraum?
    If ((X - DayArea.X) Mod (LineH + 1)) = LineH Then Exit Sub
    If ((Y - DayArea.Y) Mod (LineH + 1)) = LineH Then Exit Sub
    'sonst DownDay vergleichen
    i = (X - DayArea.X) \ (LineH + 1)
    If Not (i = DayDownX) Then Exit Sub
    i = (Y - DayArea.Y) \ (LineH + 1)
    If Not (i = DayDownY) Then Exit Sub
    'Tag geklickt -> Tag berechnen
    aDate = DateAdd("ww", DayDownX, tDate)
    aDate = DateAdd("d", DayDownY, aDate)
    RefreshDate
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
  If (UnloadMode = vbFormControlMenu) Then
    Cancel = True: Hide
  End If
End Sub
