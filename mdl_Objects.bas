Attribute VB_Name = "mdl_Objects"
Option Explicit

'''''''''''''''''''''''''''
''' OBJECT SIZING
''' Dropdown adjust
'
'Const CONTROL_OPTIONS = "Height;Left;Locked;Placement;Top;Width" 'some potentially useful settings to store and sustain
'
'Function refreshControlsOnSheet(sh As Object) 'routine enumerates all objects on the worksheet (sh), determines which have stored settings, then refreshes those settings from storage (in the defined names arena)
'
'Dim myControl As OLEObject
'Dim sBuildControlName As String
'Dim sControlSettings As Variant
'
'For Each myControl In ActiveSheet.OLEObjects
'    sBuildControlName = "_" & myControl.name & "_Range" 'builds a range name based on the control name
'    'test for existance of previously-saved settings
'    On Error Resume Next
'    sControlSettings = Evaluate(sBuildControlName) 'ActiveWorkbook.Names(sBuildControlName).RefersTo 'load the array of settings
'    If Err.Number = 0 Then ' the settings for this control are in storage, so refresh settings for the control
'        myControl.Height = sControlSettings(1)
'        myControl.Left = sControlSettings(2)
'        myControl.Locked = sControlSettings(3)
'        myControl.Placement = sControlSettings(4)
'        myControl.Top = sControlSettings(5)
'        myControl.Width = sControlSettings(6)
'    End If
'    Err.Clear
'    On Error GoTo 0
'Next myControl
'End Function
'
'Private Sub storeControlSettings(sControl As String)
'Dim sBuildControlName As String
'Dim sControlSettings(1 To 6) As Variant ' set to the number of control settings to be stored
'Dim oControl As Variant
'
'Set oControl = ActiveSheet.OLEObjects(sControl)
'
''store the settings to retain, so they can be reset on demand, thus avoiding Excel's resizing "problem"
''create array of settings to be stored, with order dictated by CONTROL_OPTIONS for consistency/documentation
'
'sControlSettings(1) = oControl.Height
'sControlSettings(2) = oControl.Left
'sControlSettings(3) = oControl.Locked
'sControlSettings(4) = oControl.Placement
'sControlSettings(5) = oControl.Top
'sControlSettings(6) = oControl.Width
'
'
'sBuildControlName = "_" & sControl & "_Range" 'builds a range name based on the control name
'
'Application.Names.Add name:="'" & ActiveSheet.name & "'!" & sBuildControlName, RefersTo:=sControlSettings, Visible:=False 'Adds the control's settings to the defined names area and hides the range name
'End Sub
'
'
'Public Sub setControlsOnSheet()
'Dim myControl As OLEObject
'
'If vbYes = MsgBox("If you click 'Yes' the settings for all controls on your active worksheet will be stored as they CURRENTLY exist. " & vbCrLf & vbCrLf _
'                & "Are you sure you want to continue (any previous settings will be overwritten)?", vbYesNo, "Store Control Settings") Then
'
'    For Each myControl In ActiveSheet.OLEObjects 'theoretically, one could manage settings for all controls of this type...
'        storeControlSettings (myControl.name)
'    Next myControl
'
'    MsgBox "Settings have have been stored", vbOKOnly
'End If
'Application.EnableEvents = True 'to ensure we're set to "fire" on worksheet changes
'End Sub
