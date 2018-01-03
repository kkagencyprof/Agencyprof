Attribute VB_Name = "Module1"
Option Explicit
Declare Sub InitCommonControls Lib "comctl32.dll" ()

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

Public Type POINTAPI   ' pt
  X As Long
  Y As Long
End Type

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Public Const WM_USER = &H400

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
                            (ByVal dwExStyle As Long, ByVal lpClassName As String, _
                             ByVal lpWindowName As String, ByVal dwStyle As Long, _
                             ByVal X As Long, ByVal Y As Long, _
                             ByVal nWidth As Long, ByVal nHeight As Long, _
                             ByVal hwndParent As Long, ByVal hMenu As Long, _
                             ByVal hInstance As Long, lpParam As Any) As Long

Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
'

' Returns the low-order word from the given 32-bit value.

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

' Returns the larger of the two passed params

Public Function Max(param1 As Long, param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else Max = param2
End Function

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would rtn a zero length string ("").
    GetStrFromBufferA = szA
  End If
End Function

Public Function XShell( _
    ByVal PathName As String, _
    Optional ByVal WindowStyle As VbAppWinStyle = vbMinimizedFocus, _
    Optional ByVal Events As Boolean = True _
  ) As Long

  'Deklarationen:
  Const PACTIVE = &H103&
  Const PQINF = &H400&
  Dim ProcId As Long
  Dim ProcHnd As Long

  'Prozess-Handle holen:
  ProcId = Shell(PathName, WindowStyle)
  ProcHnd = OpenProcess(PQINF, True, ProcId)

  'Auf Prozess-Ende warten:
  Do
    If Events Then DoEvents
    GetExitCodeProcess ProcHnd, XShell
  Loop While XShell = PACTIVE

  'Aufräumen:
  CloseHandle ProcHnd

End Function


