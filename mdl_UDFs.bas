Attribute VB_Name = "mdl_UDFs"
''BACKUP OF b_MyFunctions addin

'Option Explicit
'
'Public Function p_GetColor(Cell As Range, Optional OfText As Boolean = False)
'
''Application.Volatile True
'
'Dim CI As Double
'
'If OfText = True Then
'    CI = Cell.Font.Color
'Else
'    CI = Cell.Interior.Color
'End If
'
'p_GetColor = CI
'
'End Function
'
'Function p_TrueRand(Lower As Variant, Upper As Variant)
'    p_TrueRand = Rnd * (Upper - Lower) + Lower
'End Function
'
''Public Function p_ModDate() As String
''    If ActiveWorkbook.Path = "" Then
''        p_ModDate = ActiveWorkbook.BuiltinDocumentProperties("Creation Date")
''    Else
''        p_ModDate = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
''    End If
''End Function
'
''Public Function p_ModDateStr() As String
''    If ActiveWorkbook.Path = "" Then
''        p_ModDateStr = "Last Modified on " & Format(ActiveWorkbook.BuiltinDocumentProperties("Creation Date"), "m/d/yy hh:mm AM/PM")
''    Else
''        p_ModDateStr = "Last Modified on " & Format(ActiveWorkbook.BuiltinDocumentProperties("Last Save Time"), "m/d/yy hh:mm AM/PM")
''    End If
''End Function
'
'Function p_ExtractElement(str, n, Delimiter, Optional Clean As Boolean = False, Optional nElements As Boolean = False)
'    '=ExtractElement(text, element/column number, delimiter string)
'    '   Returns the nth element from a string,
'    '   using a specified separator character
'    Dim x As Variant
'    x = Split(str, Delimiter)
'    If n > 0 And n - 1 <= UBound(x) Then
'        If Clean = True Then
'            p_ExtractElement = Replace(x(n - 1), " ", "")
'        Else
'            p_ExtractElement = x(n - 1)
'        End If
'    Else
'        p_ExtractElement = ""
'    End If
'
'    If nElements = True Then p_ExtractElement = UBound(x) + 1
'End Function
'
'
'Function p_GetComment(incell) As String
''   Get comments from cell
'    On Error Resume Next
'    p_GetComment = incell.Comment.Text
'
'End Function
'
'
''   Get comments from cell, removes the commenters name
'Function p_GetCommentClean(incell) As String
'
'    Dim str1 As String
'    Dim var1 As Variant
'
'    On Error Resume Next
'
'    str1 = incell.Comment.Text
'
'    '   Remove commenter name
'    var1 = InStr(1, str1, ":")
'    var1 = var1 + 1
'    str1 = Mid(str1, var1, Len(str1))
'
'    '   Replace line breaks with regular spaces
'    str1 = Replace(str1, Chr(13), " ")
'    str1 = Replace(str1, Chr(10), " ")
'    '   Replace double spaces with single spaces
'    str1 = Replace(str1, "  ", " ")
'    '   Remove first character if space
'    Do While Mid(str1, 1, 1) = " "
'        str1 = Mid(str1, 2, Len(str1) - 1)
'    Loop
'
'    p_GetCommentClean = str1
'
'End Function
'
''   If formula meets criteria then something, else the value of the formula
''       Useful so you do not need to repeat the formula twice
''       Currently only equals / "=" criteria works, need to add </>
'Public Function p_IFVAR(checkcell As Variant, checkcond As String, notb As Variant) As Variant
'    If checkcell = checkcond Then
'        p_IFVAR = notb
'    Else
'        p_IFVAR = checkcell
'    End If
'End Function
'
'
'Function p_ConcatenateIF(CriteriaRange As Range, Condition As Variant, ConcatenateRange As Range, Optional Separator As String = ",") As Variant
'    'Update 20150414
'    Dim xResult As String
'    Dim i As Integer
'    On Error Resume Next
'    If CriteriaRange.Count <> ConcatenateRange.Count Then
'        ConcatenateIF_p = CVErr(xlErrRef)
'        Exit Function
'    End If
'    For i = 1 To CriteriaRange.Count
'        If CriteriaRange.Cells(i).Value = Condition Then
'            xResult = xResult & Separator & ConcatenateRange.Cells(i).Value
'        End If
'    Next i
'    If xResult <> "" Then
'        xResult = VBA.Mid(xResult, VBA.Len(Separator) + 1)
'    End If
'    p_ConcatenateIF = xResult
'    Exit Function
'End Function
'
'
'
''
''Public Function MinMaxIFS(MinMax As Byte, r1 As Range, ParamArray OtherArgs())
''
''Dim iLoop As Integer
''Dim vMy_Array As Variant
''Dim Criteria As Variant
''
''On Error GoTo My_End:
''
'''error checking on function call
''If MinMax < 1 Or MinMax > 2 Then
''    MinMaxIFS = 0: Exit Function
''End If
''If r1 Is Nothing Then 'empty range
''    MinMaxIFS = 0: Exit Function
''End If
''If UBound(OtherArgs) = -1 Then  'no criteria used - caluclate the Min of the range
''    MinMaxIFS = IIf(MinMax = 1, WorksheetFunction.Min(r1), WorksheetFunction.Max(r1))
''    Exit Function
''End If
''If ((UBound(OtherArgs) - LBound(OtherArgs)) Mod 2) = 0 Then 'uneven number of parameters - means one of them are missing
''    MinMaxIFS = 0: Exit Function
''End If
'''this is a check to make sure that there is a valid criteria RANGE (i.e. size /shape) in every second parameter
''For iLoop = LBound(OtherArgs) To UBound(OtherArgs) Step 2
''    If (OtherArgs(iLoop).Rows.Count <> r1.Rows.Count) Or _
''        (OtherArgs(iLoop).Columns.Count <> r1.Columns.Count) Then
''        MinMaxIFS = 0: Exit Function
''    End If
''Next
''
''Criteria = OtherArgs
''
'''applies the criteria (ranges and values) defined in otherargs and returns an array of values from r1 that meet that criteria
''vMy_Array = Apply_Criteria_IFS(r1, Criteria)
''
''MinMaxIFS = IIf(MinMax = 1, WorksheetFunction.Min(vMy_Array), WorksheetFunction.Max(vMy_Array))
''
''
''Exit Function
''My_End:
''MinMaxIFS = 0
''
''End Function
''
''Public Function SmallLargeIFS(Small_Large As Byte, r1 As Range, Nth As Integer, ParamArray OtherArgs())
''
''Dim iLoop As Integer
''Dim vMy_Array As Variant
''Dim Criteria As Variant
''
''On Error GoTo My_End:
''
'''error checking on function call
''If Small_Large < 1 Or Small_Large > 2 Then
''    SmallLargeIFS = 0: Exit Function
''End If
''If r1 Is Nothing Then 'empty range
''    SmallLargeIFS = 0: Exit Function
''End If
''If UBound(OtherArgs) = -1 Then 'no criteria used - caluclate the Min of the range
''    SmallLargeIFS = IIf(Small_Large = 1, WorksheetFunction.Small(r1, Nth), WorksheetFunction.Large(r1, Nth))
''    Exit Function
''End If
''If ((UBound(OtherArgs) - LBound(OtherArgs)) Mod 2) = 0 Then 'uneven number of parameters - means one of them are missing
''    SmallLargeIFS = 0: Exit Function
''End If
'''this is a check to make sure that there is a valid criteria RANGE (i.e. size /shape) in every second parameter
''For iLoop = LBound(OtherArgs) To UBound(OtherArgs) Step 2
''    If (OtherArgs(iLoop).Rows.Count <> r1.Rows.Count) Or _
''        (OtherArgs(iLoop).Columns.Count <> r1.Columns.Count) Then
''        SmallLargeIFS = 0: Exit Function
''    End If
''Next
''
''Criteria = OtherArgs
''
'''applies the criteria (ranges and values) defined in otherargs and returns an array of values from r1 that meet that criteria
''vMy_Array = Apply_Criteria_IFS(r1, Criteria)
''
''SmallLargeIFS = IIf(Small_Large = 1, WorksheetFunction.Small(vMy_Array, Nth), WorksheetFunction.Large(vMy_Array, Nth))
''
''Exit Function
''My_End:
''SmallLargeIFS = 0
''
''End Function
'
''
''
''Private Function Apply_Criteria_IFS(r1, Criteria)
''
'''applies the criteria (ranges and values) defined in otherargs and returns an array of values from r1 that meet that criteria
''
''Dim lLoop As Long
''Dim My_Array As Variant
''Dim All_Found As Boolean
''Dim vCrit_Value As Variant
''Dim iCrit_Count As Integer
''
''ReDim My_Array(0)
''
''For lLoop = 1 To r1.Rows.Count 'process all rows in range
''    All_Found = False 'control variable for monitoring if all criteria are matching on a row.
''    For iCrit_Count = LBound(Criteria) To UBound(Criteria) Step 2 'process each criteria on each row
''        Select Case Left(Criteria(iCrit_Count + 1), 2) 'identify the criteria value to match in the range
''            Case Is = "<=": vCrit_Value = Val(Mid(Criteria(iCrit_Count + 1), 3) + 1)
''            Case Is = ">=": vCrit_Value = Val(Mid(Criteria(iCrit_Count + 1), 3) - 1)
''            Case Is = "<>": vCrit_Value = UCase(Mid(Criteria(iCrit_Count + 1), 3))
''            Case Else
''                If Criteria(iCrit_Count + 1) Like "[<>=]*" Then
''                    vCrit_Value = IIf(IsNumeric(Mid(Criteria(iCrit_Count + 1), 2)), Val(Mid(Criteria(iCrit_Count + 1), 2)), UCase(CStr(Mid(Criteria(iCrit_Count + 1), 2))))
''                Else
''                    vCrit_Value = UCase(Criteria(iCrit_Count + 1))
''                End If
''        End Select
''
''        With Criteria(iCrit_Count).Parent
''            If Left(Criteria(iCrit_Count + 1), 2) = "<>" Then
''                All_Found = UCase(.Cells(lLoop, Criteria(iCrit_Count).Column)) <> vCrit_Value
''            Else
''                Select Case Left(Criteria(iCrit_Count + 1), 1)
''                    Case Is = "<": All_Found = .Cells(lLoop, Criteria(iCrit_Count).Column) < vCrit_Value  'also covers <=
''                    Case Is = ">": All_Found = .Cells(lLoop, Criteria(iCrit_Count).Column) > vCrit_Value  'also covers >=
''                    Case Else:     All_Found = IIf(IsNumeric(.Cells(lLoop, Criteria(iCrit_Count).Column)), Val((.Cells(lLoop, Criteria(iCrit_Count).Column))) = vCrit_Value, UCase(CStr((.Cells(lLoop, Criteria(iCrit_Count).Column)))) = vCrit_Value)  'covers = and no equal sign
''                End Select
''            End If
''        End With
''        If Not All_Found Then Exit For 'if any one of the search items isnt found on that row in the database , then exit the loop
''    Next
''    If All_Found Then 'store number of all search items found
''        With Criteria(iCrit_Count - 2).Parent
''            If IsNumeric(.Cells(lLoop, r1.Column)) Then 'make sure a value is available before adding it
''                My_Array(UBound(My_Array)) = .Cells(lLoop, r1.Column).Value
''                ReDim Preserve My_Array(UBound(My_Array) + 1)
''            End If
''        End With
''    End If
''Next
''
''ReDim Preserve My_Array(UBound(My_Array) - 1) 'remove last array item (it will be blank anyway)
''
''Apply_Criteria_IFS = My_Array
''
''End Function
''
''
'
'
'
'
'
'
