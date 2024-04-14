Attribute VB_Name = "mdl_5_VBE"
Option Explicit

    ' DEPENDENCIES
    ' Requires VB reference to Microsoft Extensibility Library
    ' Requires Excel Security settings set = Trust VBA Object Model

Sub Init_ExportMacros()
Dim strPath As String
strPath = Environ$("USERPROFILE")
strPath = strPath & "\OneDrive\Documents\Backups\Excel\VBA\Modules\"
'strPath = strPath & "\Documents\Excel\VBA\Modules\"
Call ExportMacros(strPath)
End Sub

Sub Init_ExportPersonalWorkbook()
Dim strPath As String
strPath = Environ$("USERPROFILE")
strPath = strPath & "\OneDrive\Documents\Backups\Excel\VBA\"
'strPath = strPath & "\Documents\Excel\VBA\Modules\"
Call ExportPersonalWorkbook(strPath)
End Sub


Private Sub ExportMacros(strPath As String)
    
    ' DEPENDENCIES
    ' Requires VB reference to Microsoft Extensibility Library
    ' Requires Excel Security settings set = Trust VBA Object Model
    
    Dim objVBProj As VBProject
    Dim objVBComp As VBComponent
    Dim i As Integer
    On Error GoTo errHandler
    
    Set objVBProj = Application.VBE.ActiveVBProject
    i = 0
    
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    If Dir(strPath, vbDirectory) = vbNullString Then GoTo errDirectory
    'If Dir(strPath) = "" Then GoTo errDirectory
    
    Debug.Print "Exporting to: " & strPath
    
    For Each objVBComp In objVBProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            objVBComp.Export strPath & objVBComp.Name & ".bas"
            Debug.Print "Exported: " & objVBComp.Name
            i = i + 1
        End If
    Next
    
    Debug.Print "-------------------------------------------"
    Debug.Print "Total Modules Exported: " & i
    
    Set objVBProj = Nothing
    
    Exit Sub

errHandler:
    Select Case Err.Number
        Case 50035
            Debug.Print "EXPORT ERROR: Could not find directory"
            MsgBox "Could not find export directory!", vbCritical, "EXPORT ERROR"
            Set objVBProj = Nothing
        Case Else
UnknownError:
            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
            Set objVBProj = Nothing
    End Select
Exit Sub

errDirectory:
    Debug.Print "EXPORT ERROR: Could not find directory"
    MsgBox "Error exporting VBA Modules. Could not find export directory!", vbCritical, "Error!"
    Set objVBProj = Nothing

End Sub


Private Sub ExportPersonalWorkbook(strPath As String)
        
    Dim strFile As String
    
    On Error GoTo errHandler

    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    If Dir(strPath, vbDirectory) = vbNullString Then GoTo errDirectory

    strFile = ThisWorkbook.Name
    ThisWorkbook.SaveCopyAs Filename:=strPath & strFile
    
    Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
UnknownError:
            MsgBox Err.Number & vbCrLf & Err.Description & vbNewLine & vbNewLine & "Please contact the developer", vbCritical, "Error!"
    End Select
Exit Sub

errDirectory:
    Debug.Print "EXPORT ERROR: Could not find directory"
    MsgBox "Error exporting Personal VBA workbook. Could not find export directory!", vbCritical, "Error!"
    
End Sub

''''    How to add Comment and Uncomment to shortcuts
'Right-click on the toolbar and select Customize...
'Select the Commands tab.
'Under Categories click on Edit, then select Comment Block in the Commands listbox.
'Drag the Comment Block entry onto the Menu Bar (yep! the menu bar)
'Note: You should now see a new icon on the menu bar.
'Make sure that the new icon is highlighted (it will have a black square around it) then
'click Modify Selection button on the Customize dialog box.
'An interesting menu will popup.
'Under name, add an ampersand (&) to the beginning of the entry.
'So now instead of "Comment Block" it should read &Comment Block.
'Press Enter to save the change.
'Click on Modify Selection again and select Image and Text.
'Dismiss the Customize dialog box.
'Highlight any block of code and press Alt-C. Voila.
'Do the same thing for the Uncomment Block or
'any other commands that you find yourself using often.

