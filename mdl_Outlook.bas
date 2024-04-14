Attribute VB_Name = "mdl_Outlook"
Option Explicit



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       OUTLOOK EVENT/CALENDAR FUNCTIONS

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          CreateOutlookEvents
'Description:   Creates events in outlook with the data from the given array
'Parameter1:    Array input
    'Fields:    Subject, Date, & Body
'Misc:          Is limited in that array fields must be manually to the outlook fields
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CreateOutlookEvents(arrVar As Variant)

Dim i As Integer
Dim x As Integer
Dim objOut As Object
Dim olNS As Outlook.Namespace
Dim objItem As Object
Dim olFolder As Outlook.Folder
Dim objAppt As Outlook.AppointmentItem

Set objOut = CreateObject("Outlook.Application")
Set olNS = objOut.GetNamespace("MAPI")
Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)

Dim strSubject As String
Dim strDate As String
Dim strBody As String
Dim strApptID As String
Dim strStartTime As String
Dim arrVarMap As Variant
Dim m As Integer
Dim strList As String

If DebugMode = False Then On Error GoTo errHandler

strStartTime = strDefaultEventStartTime
arrVarMap = TableColumnsMapping
x = 0
strList = ""
strList = "Events Created:"

'   For each record in array
For i = 1 To UBound(arrVar, 2)
    Set objItem = olFolder.Items.Add(olAppointmentItem)
    'Reset outlook fields
    strSubject = ""
    strDate = ""
    strBody = ""
    '' For each outlook field, set if exists in mapping table
    For m = LBound(arrVarMap, 1) To UBound(arrVarMap, 1)
        Select Case arrVarMap(m)
            Case "Subject"
                strSubject = arrVar(m + 1, i)
            Case "Start"
                strDate = arrVar(m + 1, i)
            Case "Body"
                strBody = arrVar(m + 1, i)
        End Select
    Next m
    
    'Add footer to event body
    strBody = strBody & vbNewLine & vbNewLine & strOutlookFooter
    strBody = strBody & vbNewLine & strUniqueOutlookID

    'Check if date field is valid
    If IsDate(strDate) = False Then
        MsgBox "The event " & Chr(34) & strSubject & Chr(34) & " contains an invalid date field. This event will NOT be created." & vbNewLine & vbNewLine & _
            "Please fix this value in the worksheet and try again", vbCritical, "Invalid Date"
        GoTo SkipEvent
    End If

    ' Create Outlook appointment
    With objItem
        '.Start = CDate(strDate) + TimeValue(strStartTime)
        .Start = DateValue(strDate) + TimeValue(strStartTime)
        .Duration = 60
        .Subject = strSubject
        .Body = strBody
        strApptID = .GlobalAppointmentID
        .Save
    End With

    ' Success counter
    x = x + 1
    strList = strList & vbNewLine & strSubject

SkipEvent:
Next i

MsgBox "Successfully created " & x & " of " & i - 1 & " new Outlook events." & vbNewLine & vbNewLine & strList

Set objItem = Nothing
Set objOut = Nothing

Exit Sub

errHandler:
Select Case Err.Number
    Case Else
UnknownError:
        Set objItem = Nothing
        Set objOut = Nothing
        MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "There was an unknown error in the process", vbCritical, "Error!"
End Select

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          GetOutlookEvents
'Description:   Stores outlook event data into an array. Event description must contain UniqueID (public const)
'Output:        Array
'Misc:          Appends UniqueID at end, may need to remove before passing on
'Contingencies
'   Subs:
'   Functions:
'   Variables:  strUniqueOutlookID, strDefaultEventStartTime
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetOutlookEvents()

Dim objOut As Object
Dim olNS As Outlook.Namespace
Dim objItem As Object
Dim olFolder As Outlook.Folder
Dim objAppt As Outlook.AppointmentItem
Dim strSearch As String
Dim i As Integer 'counter
Dim m As Integer 'counter
Dim arrVar As Variant
Dim arrVarMap As Variant

Set objOut = CreateObject("Outlook.Application")
Set olNS = objOut.GetNamespace("MAPI")
Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)

strSearch = strUniqueOutlookID
i = 0
ReDim arrVar(1 To 3, 1 To 1)

arrVarMap = TableColumnsMapping

For Each objAppt In olFolder.Items
    If InStr(objAppt.Body, strSearch) > 0 Then
        i = i + 1
        ReDim Preserve arrVar(1 To UBound(arrVar, 1), 1 To i)
        
        For m = LBound(arrVarMap, 1) To UBound(arrVarMap, 1)
            Select Case arrVarMap(m)
                Case "Subject"
                    arrVar(m + 1, i) = objAppt.Subject
                Case "Start"
                    arrVar(m + 1, i) = objAppt.Start
                Case "Body"
                    arrVar(m + 1, i) = objAppt.Body
            End Select
        Next m
    End If
Next objAppt

Set objOut = Nothing
Set objItem = Nothing

If i > 0 Then GetOutlookEvents = arrVar

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          DeleteOutlookEvents
'Description:   Deletes Outlook events that match the array data
'Parameter1:    Array input
    'Fields:    Subject, Date, & Body
'Misc:          Outlook events must also have the Unique ID in the body section
'Contingencies: strUniqueOutlookID
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub DeleteOutlookEvents(arrVar As Variant)

Dim objOut As Object
Dim olNS As Outlook.Namespace
Dim objItem As Object
Dim olFolder As Outlook.Folder
Dim objAppt As Outlook.AppointmentItem

Dim strSubject As String
Dim strDate As String
Dim strBody As String
Dim strApptID As String
Dim strStartTime As String
Dim strSearch As String

Dim i As Integer    ' array item counter
Dim f As Integer    ' array item count
Dim a As Integer    ' appointment item counter
Dim x As Integer    ' success counter

If DebugMode = False Then On Error GoTo errHandler

Set objOut = CreateObject("Outlook.Application")
Set olNS = objOut.GetNamespace("MAPI")
Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)

strSearch = strUniqueOutlookID

'   Loop through every appointment in outlook
'   If appt body contains unique id then check if the appt subject/date match the array data
'   If so, delete
x = 0
f = UBound(arrVar, 2) - LBound(arrVar, 2) + 1

For a = olFolder.Items.Count To 1 Step -1
    Set objItem = olFolder.Items(a)
    If objItem.Class = olAppointment Then
        If InStr(objItem.Body, strSearch) > 0 Then
            For i = 1 To f
                If objItem.Subject = arrVar(1, i) And Int(objItem.Start) = Int(CDate(arrVar(2, i))) Then
                    objItem.Delete
                    x = x + 1
                    GoTo NextItem:
                End If
            Next i
        End If
    End If
NextItem:
Next a

Set objItem = Nothing
Set objOut = Nothing

MsgBox "Successfully deleted " & x & " of " & f & " Outlook appointments."

Exit Sub

errHandler:
Select Case Err.Number
    Case Else
UnknownError:
        Set objItem = Nothing
        Set objOut = Nothing
        MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "There was an unknown error in the process", vbCritical, "Error!"
End Select

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          CheckOutlookEventExists
'Description:   Checks if any Outlook events already exist compared to the data in the given array, if exists, then remove from array
'Parameter1:    Array input
    'Fields:
'Misc:          Outlook events must also have the Unique ID in the body section
'Contingencies: strUniqueOutlookID
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CheckOutlookEventExists(arrVar As Variant)

Dim i As Integer
Dim x As Integer
Dim a As Integer
Dim m As Integer

Dim objOut As Object
Dim olNS As Outlook.Namespace
Dim objItem As Object
Dim olFolder As Outlook.Folder
Dim objAppt As Outlook.AppointmentItem

Dim strSubject As String
Dim strDate As String
Dim intSubject As Integer 'Location of Subject field in array
Dim intDate As Integer      'Location of Date field in array
Dim strBody As String
Dim strApptID As String
Dim strStartTime As String
Dim strSearch As String
Dim arrVarMap As Variant
Dim arrOutlookEvents As Variant
Dim NoEventsFlag As Boolean

If DebugMode = False Then On Error GoTo errHandler

Set objOut = CreateObject("Outlook.Application")
Set olNS = objOut.GetNamespace("MAPI")
Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)

strSearch = strUniqueOutlookID

'   Loop through every appointment in outlook
'   If appt body contains unique id then check if the appt subject/date match the array data
'   If so, delete
x = 0
'For Each objAppt In olFolder.Items

arrVarMap = TableColumnsMapping
arrOutlookEvents = GetOutlookEvents
If IsVarArrayEmpty(arrOutlookEvents) = True Then NoEventsFlag = True

'   For each record in array
For i = UBound(arrVar, 2) To LBound(arrVar, 2) Step -1
    'Set objItem = olFolder.Items.Add(olAppointmentItem)
    'Reset outlook fields
    strSubject = ""
    strDate = ""
    'strBody = ""
    
    '' For each outlook field, set if exists in mapping table
    For m = LBound(arrVarMap, 1) To UBound(arrVarMap, 1)
        Select Case arrVarMap(m)
            Case "Subject"
                strSubject = arrVar(m + 1, i)
                intSubject = m + 1
            Case "Start"
                strDate = arrVar(m + 1, i)
                intDate = m + 1
'            Case "Body"
'                strBody = arrVar(m + 1, i)
        End Select
    Next m
    
    'Check if date field is valid
    If IsDate(strDate) = False Then
        GoTo NextItem
    End If
    
    If NoEventsFlag = False Then
        '   For each event in outlook, if match then delete
        '   Criteria is Name AND Date must match
        For a = UBound(arrOutlookEvents, 2) To LBound(arrOutlookEvents, 2) Step -1
            If arrOutlookEvents(intSubject, a) = strSubject And Int(arrOutlookEvents(intDate, a)) = strDate Then
                CheckOutlookEventExists = DeleteArrayElementAt(arrVar, i)
                GoTo NextItem:
            End If
        Next a
    End If
    
NextItem:
Next i

CheckOutlookEventExists = arrVar

Set objItem = Nothing
Set objOut = Nothing

Exit Function

errHandler:
Select Case Err.Number
    Case Else
UnknownError:
        Set objItem = Nothing
        Set objOut = Nothing
        MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "There was an unknown error in the process", vbCritical, "Error!"
End Select

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       OLD OUTLOOK FUNCTIONS

'Sub ListAppointments()
'
'    Dim olApp As Object
'    Dim olNS As Object
'    Dim olFolder As Object
'    Dim olApt As Object
'    Dim NextRow As Long
'
'    Set olApp = CreateObject("Outlook.Application")
'
'    Set olNS = olApp.GetNamespace("MAPI")
'
'    Set olFolder = olNS.GetDefaultFolder(9) 'olFolderCalendar
'
'    Range("A1:D1").Value = Array("Subject", "Start", "End", "Location")
'
'    NextRow = 2
'
'    For Each olApt In olFolder.Items
'        Cells(NextRow, "A").Value = olApt.Subject
'        Cells(NextRow, "B").Value = olApt.Start
'        Cells(NextRow, "C").Value = olApt.End
'        Cells(NextRow, "D").Value = olApt.Location
'        NextRow = NextRow + 1
'    Next olApt
'
'    Set olApt = Nothing
'    Set olFolder = Nothing
'    Set olNS = Nothing
'    Set olApp = Nothing
'
'    Columns.AutoFit
'
'End Sub

'Sub MailDraftATS()
'
'Dim objOut As Object
'Dim objMail As Object
'Dim strTo As String
'Dim strSubject As String
'Dim strBody As String
'Dim strRepFolder
'Dim strRepName
'Dim strLocation As String
'
'Set objOut = CreateObject("Outlook.Application")
'Set objMail = objOut.CreateItem(0)
'
''   Email info
''strTo = "jhendricks@gettaxcredits.com"
'strTo = "ats@gettaxcredits.com"
'strSubject = "ATS Report" & " - " & WeekdayName(Weekday(Date)) & " " & Date
'strRepFolder = "<a href= '\\STCE01\Public\Inventory'> Z:\Inventory </a>"
'strBody = _
'    "<P STYLE='font-family:Calibri;font-size:11pt;margin-bottom:.0001pt'> Attached is the current Available To Sell Report. Please let me know if there are any questions." & _
'    "<P STYLE='font-family:Calibri;font-size:10pt;color:gray;margin-bottom:.0001pt'><i>To stay up to date with the latest version of this report, please visit: </i>" & strRepFolder
'
''   The name of the report file that will be attached
''strRepName = "2016_Available_To_Sell"
'strRepName = Range("A2").Value
'strRepName = Replace(strRepName, ":", "")
'strRepName = strRepName & Format(Now(), "_YYMMDD")
'strRepName = strRepName & ".pdf"
'strRepName = Replace(strRepName, " ", "_")
'
''   The location of the above report file to be attached
'strLocation = "S:\Finance\Analytics and Modeling\Credit Inventory\" & strRepName
'
''   Check if report file (attachment) exists, dialog box if it doesn't
'If Dir(strLocation) = "" Then
'    MsgBox "Excel cannot find the report file in: " & vbNewLine & strLocation & _
'            vbNewLine & vbNewLine & "Please manually attach the report file to the email draft.", _
'            vbExclamation, "ERROR: File Not Found"
'    'GoTo errHandler
'End If
'
'On Error Resume Next
'
''   Create new email draft
'With objMail
'    .Display
'    .To = strTo
'    .Subject = strSubject
'    '.HTMLBody = strBody & "<br>" & .HTMLBody
'    .HTMLBody = strBody & .HTMLBody                 ' this insures that the signature is still added below body of email
'    .Attachments.Add (strLocation)
'End With
'
'errHandler:
'
'Set objMail = Nothing
'Set objOut = Nothing
'
'End Sub

