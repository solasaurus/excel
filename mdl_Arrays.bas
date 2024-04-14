Attribute VB_Name = "mdl_Arrays"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       STANDARD ARRAY FUNCTIONS


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          GetTableData
'Description:   Stores table data into an array. Only collects data from given columns
'Parameter1:    Fields / Column names
'Output:        Array
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetTableData(Fields As Variant) ' add

Dim ws As Worksheet
Dim tbl As ListObject
Dim r As Integer        ' number of rows
Dim ri As Integer       ' row counter
Dim f As Integer        ' number of fields
Dim fi As Integer       ' field counter
Dim arrItems As Variant
Dim x As Integer        ' successful loop counter
Dim i As Integer
Dim strVar As String

If DebugMode = False Then On Error GoTo errHandler

'   Number of fields (table columns)
f = UBound(Fields, 1) - LBound(Fields, 1) + 1

Set ws = ActiveSheet
Set tbl = ws.ListObjects(strTableName)
r = tbl.ListRows.Count

'   Setup initial array
ReDim arrItems(1 To f, 0 To 0)
x = 0

For ri = 1 To r                                                     ' Loop through each row in table
        x = x + 1                                                   ' Succesful item counter
        ReDim Preserve arrItems(1 To f, 1 To x)                     ' dynamic array grows in size for every row that matches criteria
        For fi = 1 To f                                             ' Loop through each field
            arrItems(fi, x) = tbl.DataBodyRange.Cells(ri, tbl.ListColumns(Fields(fi - 1)).Index).Value  ' Load array with table data
        Next fi
Next ri

GetTableData = arrItems

Exit Function

errHandler:
Select Case Err.Number

    Case 9 ' Subscript out of range / Can't find table/columns error
        ' Store table and column data in single string for error message
        For i = LBound(Fields, 1) To UBound(Fields, 1)
            If i = LBound(Fields, 1) Then
                strVar = Chr(34) & Fields(i) & Chr(34)
            Else
                strVar = strVar & ", " & Fields(i)
            End If
        Next i
            MsgBox "Could not find the table data. Please check that the below information is accurate in the worksheet:" & vbNewLine & vbNewLine & _
                "Worksheet Name: " & ws.Name & vbNewLine & _
                "Table Name: " & strTableName & vbNewLine & _
                "Table Columns: " & strVar _
                , vbCritical, "Error"
    Case Else
UnknownError:
            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
End Select

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          RunCriteriaOnArray
'Description:   Checks if array fields match given criteria
'Parameter1:    Array input
'Parameter2:    The location of the field to be tested in the array (assumes field values located in 2nd dimension of array)
'Parameter3:    The criteria value being compared against the array values
'Parameter4:    [Optional] The operator used to compare the array value to the criteria value. Default value is "="
'Output:        Array
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunCriteriaOnArray(arrVar As Variant, FieldPosition As Integer, CriteriaValue As Variant, Optional Operator As String = "=")

Dim i As Integer
Dim var1 As Variant
Dim TypeFlag As Boolean

On Error GoTo errHandler
If (TypeName(CriteriaValue) = "String") Or (Operator = "=") Then TypeFlag = True

' Loop through array and check if value matches criteria, if not delete
For i = UBound(arrVar, 2) To LBound(arrVar, 2) Step -1
    var1 = arrVar(FieldPosition, i)
    
Select Case TypeFlag
    Case True
        If Evaluate(Chr(34) & var1 & Chr(34) & Operator & Chr(34) & CriteriaValue & Chr(34)) = False Then
            ' Delete all array values at the index
            Debug.Print var1 & " Deleted"
            RunCriteriaOnArray = DeleteArrayElementAt(arrVar, i)
        Else
           Debug.Print var1
        End If
    Case False
        If Evaluate(CLng(var1) & Operator & CLng(CriteriaValue)) = False Then
            ' Delete all array values at the index
            Debug.Print var1 & " Deleted"
            RunCriteriaOnArray = DeleteArrayElementAt(arrVar, i)
        Else
           Debug.Print var1
        End If
End Select

NextVar:
Next i

Shutdown:
RunCriteriaOnArray = arrVar

Exit Function

errHandler:
Select Case Err.Number
    Case 13 ' Type Mismatch
        Resume NextVar
    Case Else
UnknownError:
            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
        GoTo Shutdown
End Select

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          RunCriteriaOnArray
'Description:   Checks if array fields match given criteria
'Parameter1:    Array input
'Parameter2:    The location of the field to be tested in the array (assumes field values located in 2nd dimension of array)
'Parameter3:    The criteria value being compared against the array values
'Parameter4:    [Optional] The operator used to compare the array value to the criteria value. Default value is "="
'Output:        Array
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunCriteriaOnDateArray(arrVar As Variant, FieldPosition As Integer, CriteriaValue As Variant, Optional Operator As String = "=")

Dim i As Integer
Dim var1 As Variant

On Error GoTo errHandler

' if operator not = "=" OR if Typename not = "String" then convert to long

' Loop through array and check if value matches criteria, if not delete
For i = UBound(arrVar, 2) To LBound(arrVar, 2) Step -1
    var1 = arrVar(FieldPosition, i)
    
    If Evaluate(CLng(var1) & Operator & CLng(CriteriaValue)) = False Then
    'If Evaluate(var1 & Operator & CriteriaValue) = False Then
    'If Evaluate(Chr(34) & arrVar(FieldPosition, i) & Chr(34) & Operator & Chr(34) & CriteriaValue & Chr(34)) = False Then
        ' Delete all array values at the index
        'Debug.Print var1 & " Deleted"
        RunCriteriaOnDateArray = DeleteArrayElementAt(arrVar, i)
'    Else
'        Debug.Print var1
    End If

NextVar:
Next i

Shutdown:
RunCriteriaOnDateArray = arrVar

Exit Function

errHandler:
Select Case Err.Number
    Case 13 ' Type Mismatch
        Resume NextVar
    Case Else
UnknownError:
            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
        GoTo Shutdown
End Select

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          DeleteArrayElementAt
'Description:   Deletes all array values at the given index
'Parameter1:    Array input
'Parameter2:    The location of the value in the (2nd dimension of the) array to be deleted
'Output:        Array
'Misc:          Currently only works on a two dimensional array. Adding GetArrayDimensions functions could allow more dimensions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function DeleteArrayElementAt(arrVar As Variant, ByVal Index As Integer)

Dim i As Integer
Dim intArr As Integer
Dim x As Integer

'intArr = GetArrayDimensions(arrVar)

'For i = 1 To intArr

' Removes array values at index, shifts everything beyond index up one place in array, index is overwritten
For i = LBound(arrVar, 1) To UBound(arrVar, 1)
    For x = (Index + 1) To UBound(arrVar, 2)
        arrVar(i, x - 1) = arrVar(i, x)
    Next x
Next i

' Removes the (now empty) last value in array
'   If there was only 1 value, then set to nothing
If LBound(arrVar, 2) > (UBound(arrVar, 2) - 1) Then
    'ReDim arrVar(0, 0)
    Erase arrVar
Else
    ReDim Preserve arrVar(LBound(arrVar, 1) To UBound(arrVar, 1), LBound(arrVar, 2) To (UBound(arrVar, 2) - 1))
End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:          IsVarArrayEmpty
'Description:   Determines if a variant array is empty
'Parameter1:    Array input
'Output:        Boolean
'Misc:          Works well after using DeleteArrayElementAt to determine if the last value was removed
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IsVarArrayEmpty(arrVar As Variant)

Dim i As Integer

On Error Resume Next
    i = UBound(arrVar, 1)
If Err.Number = 0 Then
    IsVarArrayEmpty = False
Else
    IsVarArrayEmpty = True
End If

End Function

