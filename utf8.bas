Attribute VB_Name = "Module2"
Option Explicit


Public Sub change_field_size(DBPath As String, _
  tblName As String, fldName As String, fldSize As Integer)
    ' this routine changes the field size
    
    Dim db As Database
    Dim td As TableDef
    Dim fld As Field
        
    On Error GoTo errhandler

    Set db = OpenDatabase(DBPath)
    Set td = db.TableDefs(tblName)
    
    If td.Fields(fldName).Type <> dbText Then
        ' wrong field type
        db.Close
        Exit Sub
    End If
    
    If td.Fields(fldName).Size = fldSize Then
        ' the field width is correct
        db.Close
        Exit Sub
    End If
    
    ' create a temp feild
    td.Fields.Append td.CreateField("temp", dbText, fldSize)
    td.Fields("temp").AllowZeroLength = True
    td.Fields("temp").DefaultValue = """"""

    ' copy the info into the temp field
    db.Execute "Update " & tblName & " set temp = " & fldName & " "
    
    ' delete the field
    td.Fields.Delete fldName
    
    ' rename the field
    td.Fields("temp").name = fldName
    db.Close
    
'======================================================================
Exit Sub

errhandler:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description & vbCrLf & "Change Field Size Routine", vbCritical, App.Title

End Sub

Public Function utf8sbjdecode(s$) As String
Dim ut$, ut1$, ut2$, sbj$, i%, utyp As String, c$

sbj$ = s$
utf8sbjdecode = ""
      ut1$ = ""
      ut2$ = strrepl(sbj$, "==", vbCrLf)
      For i% = 1 To linesof(ut2$)
        c$ = lineof(i%, ut2$)
        utyp = "q": If InStr(LCase(c$), "?b?") > 0 Then utyp = "b"
        Do
          ut$ = cut_d1(c$, "?"): c$ = cut_d2bis(c$, "?")
        Loop Until c$ = "=" Or c$ = ""
        If utyp = "b" Then
'Debug.Print ut$
          ut$ = Base64DecodeString(ut$)
'Debug.Print ut$
        End If
        ut$ = form1.ask_agencyprof_com(ut$)
        If ut$ <> "" Then ut1$ = ut1$ + ut$
      Next i%
utf8sbjdecode = ut1$
End Function

Public Function URLEncode(StringToEncode As String, Optional _
   UsePlusRatherThanHexForSpace As Boolean = False) As String

Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case Asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & _
              Format(Hex(Asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function


Public Function URLDecode(StringToDecode As String) As String

Dim TempAns As String
Dim CurChr As Integer

CurChr = 1

Do Until CurChr - 1 = Len(StringToDecode)
  Select Case Mid(StringToDecode, CurChr, 1)
    Case "+"
      TempAns = TempAns & " "
    Case "%"
      TempAns = TempAns & Chr(Val("&h" & _
         Mid(StringToDecode, CurChr + 1, 2)))
       CurChr = CurChr + 2
    Case Else
      TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
  End Select

CurChr = CurChr + 1
Loop

URLDecode = TempAns
End Function

