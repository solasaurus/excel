Attribute VB_Name = "mdl_Dialogs"
Option Explicit

Private Sub RangeSelectionPrompt()
    Dim rng As Range
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

    MsgBox "The cells selected were " & rng.Address
End Sub

