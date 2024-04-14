Attribute VB_Name = "mdl_3_MetaCode"
Option Explicit

Private Sub Template()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

'   Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

''' CODE GOES HERE

'   Shutdown
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

Exit Sub

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub
