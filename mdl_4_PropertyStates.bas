Attribute VB_Name = "mdl_4_PropertyStates"
Option Explicit

Private Type Callable
  o As Object
  p As String
End Type


Sub PrintCurrentStates()
Dim Pattern As String

Pattern = "!@@@@@@@@@@@@@@@@@"
Dim strProperty() As String
Dim i As Integer

'   Number of Properties
i = 4
ReDim strProperty(0 To i - 1)

'   Properties and their default values
strProperty(0) = "Application.ScreenUpdating"
strProperty(1) = "Application.EnableEvents"
strProperty(2) = "Application.Calculation"
strProperty(3) = "Application.StatusBar"

'   Get current state for each property, reset to default
For i = 0 To UBound(strProperty, 1)
'   Current state
    Debug.Print _
        "--------------------------------------------" & vbNewLine _
        ; Format("Property: ", Pattern) & strProperty(i) & vbNewLine & _
        Format("Current State: ", Pattern); GetProperty(strProperty(i))
Next i

''   Reference:
'        -4105 = xlCalculationAutomatic
'        -4135 = xlCalculationManual
'        2 = xlCalculationSemiAutomatic

End Sub

Sub Init_ResetDefaultStates()
    Call ResetDefaultStates
End Sub

Sub ResetDefaultStates(Optional bPrint As Boolean = True)
Dim Pattern As String

Pattern = "!@@@@@@@@@@@@@@@@@"
Dim strProperty() As String
Dim strDefault() As Variant
Dim i As Integer

'   Number of Properties
i = 3
ReDim strProperty(0 To i - 1)
ReDim strDefault(0 To i - 1)

'   Properties and their default values
strProperty(0) = "Application.ScreenUpdating"
strDefault(0) = True

strProperty(1) = "Application.EnableEvents"
strDefault(1) = True

strProperty(2) = "Application.Calculation"
strDefault(2) = xlCalculationAutomatic

'strProperty(3) = "Application.StatusBar"
'strDefault(3) = False

'   Get current state for each property, reset to default
For i = 0 To UBound(strProperty, 1)
'   Current state
    If bPrint = True Then Debug.Print _
        "--------------------------------------------" & vbNewLine _
        ; Format("Property: ", Pattern) & strProperty(i) & vbNewLine & _
        Format("Current State: ", Pattern); GetProperty(strProperty(i))
    
'   Set to default state
    SetProperty strProperty(i), strDefault(i)

'   New state
    If bPrint = True Then Debug.Print _
        Format("New State: ", Pattern); GetProperty(strProperty(i))
Next i

''   Reference:
'        -4105 = xlCalculationAutomatic
'        -4135 = xlCalculationManual
'        2 = xlCalculationSemiAutomatic

End Sub

Public Sub SetProperty(ByVal path As String, ByVal Value As Variant, Optional ByVal RootObject As Object = Nothing)
  With GetObjectFromPath(RootObject, path)
    If IsObject(Value) Then
      CallByName .o, .p, VbSet, Value
    Else
      CallByName .o, .p, VbLet, Value
    End If
  End With
End Sub

Public Function GetProperty(ByVal path As String, Optional ByVal RootObject As Object = Nothing) As Variant
  With GetObjectFromPath(RootObject, path)
    GetProperty = CallByName(.o, .p, VbGet)
  End With
End Function

Private Function GetObjectFromPath(ByVal RootObject As Object, ByVal path As String) As Callable
  'Returns the object that the last .property belongs to
  Dim s() As String
  Dim i As Long

  If RootObject Is Nothing Then Set RootObject = Application

  Set GetObjectFromPath.o = RootObject

  s = Split(path, ".")

  For i = LBound(s) To UBound(s) - 1
    If Len(s(i)) > 0 Then
      Set GetObjectFromPath.o = CallByName(GetObjectFromPath.o, s(i), VbGet)
    End If
  Next

  GetObjectFromPath.p = s(UBound(s))
End Function





