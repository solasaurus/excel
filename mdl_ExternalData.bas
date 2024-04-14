Attribute VB_Name = "mdl_ExternalData"
''Option Explicit
''
''Sub UpdateSalesData()
''
''Dim wb As Workbook
''Dim ws As Worksheet
''
''Dim strFolder As String
''Dim strFilePath As String
''Dim wbExt As Workbook
''Dim wsExt As Worksheet
''
''Dim tbl As ListObject
''Dim tblExt As ListObject
''
''Dim screenUpdateState As Variant
''Dim eventsState As Variant
''Dim calcState As Variant
''
''''' Startup
''On Error GoTo errHandler
''screenUpdateState = Application.ScreenUpdating
''eventsState = Application.EnableEvents
''calcState = Application.Calculation
''Application.ScreenUpdating = False
''Application.EnableEvents = False
''Application.Calculation = xlCalculationManual
''
'''strFilePath = "C:\Users\JHendricks\Documents\Finance\Credit_Sales\Actuals\2016_Principal_Div_Credit_Sales_v02_10.xlsm"
''
'''   Gets the file with the latest date modified property from the given folder, based on optional criteria
'''strFilePath = fnFindRecentFile("C:\Users\JHendricks\Documents\Finance\Credit Sales\Actuals\", "Principal")
''strFolder = "S:\Sales\2016\"
''strFilePath = fnFindRecentFile(strFolder, "Credit Sales")
''
'''   Inventory Workbook
''Set wb = ThisWorkbook
''Set ws = wb.Sheets("Credit Sales")
''Set tbl = ws.ListObjects("tbl_SALES")
''
'''   Sales Workbook
''Set wbExt = Workbooks.Open(strFilePath)
''Set wsExt = wbExt.Sheets("2016 Principal")
''Set tblExt = wsExt.ListObjects("tbl_Sales")
''
'''   Copy Sales Data
''With tblExt
''    .Range.AutoFilter
''    .ShowAutoFilter = True
''    '.Range.Copy
''    .DataBodyRange(1, 1).Resize(tblExt.ListRows.Count, 7).Copy
''End With
''
'''   Clear destination
'''   Might need to delete inventory sales data before updating
''
'''   Paste into destination
'''ws.Cells(2, 1).PasteSpecial xlPasteValues
''tbl.DataBodyRange(1, 1).Cells.PasteSpecial xlPasteValues
''
'''   Clear cutcopymode / makes sure dialog box asking to clear doesnt show up
''Application.CutCopyMode = False
''
'''   Close Sales Workbook without saving
''wbExt.Close SaveChanges:=False
''
''ws.Cells(1, 1).Select
''
'''CreateObject("WScript.Shell").PopUp "Sales Data Successfully Updated", 2
'''MsgBox "Sales Data Successfully Updated"
''
''Application.ScreenUpdating = screenUpdateState
''Application.EnableEvents = eventsState
''Application.Calculation = calcState
''
''Exit Sub
''
''errHandler:
''Application.ScreenUpdating = screenUpdateState
''Application.EnableEvents = eventsState
''Application.Calculation = calcState
''MsgBox "There was an error in the update process." & _
''        vbCrLf & "Please make sure you have a connection to the sales folder." & _
''        vbCrLf & "Location currently set to: " & strFolder, _
''        vbExclamation, "Update Unsuccesful"
''
''End Sub
''
''
''Function fnFindRecentFile(strPath As String, Optional strCriteria As String)
''
''Dim oFSO As Object
''Dim oFiles As Object
''Dim oFile As Object
''
'''Dim strPath As String
''Dim strFile As String
'''Dim strCriteria As String
''
''Dim arrVar() As Variant
''Dim j As Integer
''
'''strPath = "C:\Users\JHendricks\Documents\Finance\Credit Sales\Actuals\"
'''strCriteria = "Principal"
''
''Set oFSO = CreateObject("Scripting.FileSystemObject")
''
''Set oFiles = oFSO.GetFolder(strPath).Files
''
''ReDim arrVar(1 To 2, 0)
''j = -1
''For Each oFile In oFiles
''    If Not oFile.Name Like "~$*" Then                                               '   If not temp file / open file
''        If oFile.Name Like "*" & strCriteria & "*" Then                             '   If file name match criteria
''            j = j + 1
''            ReDim Preserve arrVar(LBound(arrVar, 1) To UBound(arrVar, 1), 0 To j)
''            arrVar(1, j) = oFile.Name
''            arrVar(2, j) = oFile.DateLastModified
''        End If
''    End If
''Next oFile
''
'''   Sort files by data last modified
''fnSort2DArray arrVar
''
'''   Set to most recent file (file with greatest date in array)
''strFile = arrVar(1, UBound(arrVar, 2))
''
''fnFindRecentFile = strPath & strFile
''
''End Function
''

