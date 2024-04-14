Attribute VB_Name = "mdl_Slicers"
Option Explicit

'Sub SlicerViewToggle(varSidePanelState As Integer)
'
'Dim wb As Workbook
'Dim ws As Worksheet
'Dim sc As SlicerCache
'Dim sl As Slicer
'
'Set wb = ActiveWorkbook
'Set ws = ActiveSheet
'
'For Each sc In wb.SlicerCaches
'    For Each sl In sc.Slicers
'        sl.Shape.Visible = IIf(varSidePanelState = 0, True, False)
'    Next sl
'Next sc
'
'End Sub
'
'
'Sub List_Slicers()
''Description: List all slicers and sheet names in the Immediate window
''Author: Jon Acampora, Excel Campus
''Source: http://www.excelcampus.com/library/vba-macro-list-all-slicers/
'
'Dim sc As SlicerCache
'Dim sl As Slicer
'
'    For Each sc In ActiveWorkbook.SlicerCaches
'        For Each sl In sc.Slicers
'            Debug.Print sl.Caption & " | " & sl.Parent.Name
'            'sl.Caption = slicer header caption
'            'sl.Parent.Name = worksheet name
'        Next sl
'    Next sc
'
'End Sub

