Attribute VB_Name = "mdl_ErrorHandling"
Option Explicit

'found on githib, need to research later

'Public Function Log(ErrNumber As Long, ErrDescription As String, Optional pLogFileName As String) As Boolean
'    On Error GoTo ErrHandler
'    Log = True
'    Dim LogFile As String
'
'    If IsMissing(pLogFileName) Or pLogFileName = vbNullString Then
'        pLogFileName = ThisWorkbook.FullName & "_debug.log"
'    End If
'    Dim FileNum
'    FileNum = FreeFile
'    Open pLogFileName For Append As #FileNum
'    Print #FileNum, Now & " ; " & Application.UserName & " ; " & ErrNumber & " ; " & ErrDescription
'    Close #FileNum
'
'    GoTo Finish:
'ErrHandler:
'    Log = False
'test
'Finish:
'End Function




''UnknownError:
''            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
''            Set objVBProj = Nothing
