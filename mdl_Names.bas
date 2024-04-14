Attribute VB_Name = "mdl_Names"
Option Explicit

'   Removes any names that contain certain value
'   BE CAREFUL YOU HAVE THE RIGHT WORKBOOK ACTIVE
Private Sub NamesRemoveRefErrors()

Dim wb As Workbook
Dim Name As Name
Dim strSearch1 As String
Dim strSearch2 As String
Dim k As Integer
Dim d As Integer
Dim Pattern As String

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

''' Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Pattern = "!@@@@@@@@@@"

strSearch1 = "#REF!"
strSearch2 = ".xls"

k = 0
d = 0

Debug.Print Format("ACTION", Pattern) & " ||  " & "Name"
Debug.Print "---------------------------------"

For Each Name In ActiveWorkbook.Names
    If InStr(Name.Value, strSearch1) > 0 Or InStr(Name.Value, strSearch2) > 0 Then
        Debug.Print Format("DELETED:", Pattern) & " ||  " & Name.Name
        d = d + 1
'''''''''''
        'Uncomment below to actually delete!
        'name.Delete
''''''''''''
    Else
        Debug.Print Format("KEPT:", Pattern) & " ||  " & Name.Name
        k = k + 1
    End If
Next Name

Pattern = "!@@@@@@@@@@@@@@@@"

Debug.Print "---------------------------------"
Debug.Print Format("TOTAL REVIEWED: ", Pattern) & k + d
Debug.Print Format("TOTAL KEPT: ", Pattern) & Format(k, "!@@@@@@") & " ||  " & Format((k / (k + d)), "Percent")
Debug.Print Format("TOTAL DELETED: ", Pattern) & Format(d, "!@@@@@@") & " ||  " & Format((d / (k + d)), "Percent")
Debug.Print "---------------------------------"
Debug.Print "---------------------------------"

Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

Exit Sub

errHandler:
MsgBox "There was an error in the process", vbExclamation
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub


Private Sub GetHiddenNames()

    Dim Name As Name
    
    For Each Name In Application.Names
        If Name.Visible = False Then
            Debug.Print Name.Name
        End If
    Next Name

End Sub


Private Sub GetAllNames()

    Dim Name As Name
    
    For Each Name In Application.Names
        Debug.Print Name.Name
    Next Name

End Sub
