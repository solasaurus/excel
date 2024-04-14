Attribute VB_Name = "mdl_Sorts"
Option Explicit

'   Simple Bubble Sort
Private Sub fnSort(a)
    For i = UBound(a) - 1 To 0 Step -1
        For j = 0 To i
            If a(j) > a(j + 1) Then
                temp = a(j + 1)
                a(j + 1) = a(j)
                a(j) = temp
            End If
        Next
    Next
End Sub

'   2 Dimensional Bubble Sort Array
Private Sub fnSort2DArray(a)
Dim i As Integer
Dim j As Integer
Dim temp As Variant

    For i = UBound(a, 2) - 1 To 0 Step -1       'Need to test and make sure this is still acurrate from orignal / 1D version
        For j = 0 To i
            If a(2, j) > a(2, j + 1) Then
                '   1D
                temp = a(2, j + 1)
                a(2, j + 1) = a(2, j)
                a(2, j) = temp
                '   2D
                temp = a(1, j + 1)
                a(1, j + 1) = a(1, j)
                a(1, j) = temp
            End If
        Next
    Next

End Sub
