' This macro function resets all formats on selected cells
Sub ResetFormatOnSelected()
    With Selection
        .ClearFormats
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
        .Hyperlinks.Delete
    End With
End Sub

' This macro function trims white spaces, tabs, and return lines from selected cells. 
Sub TrimSelectedCells()
    Application.ScreenUpdating = False
    Dim Cell As Range, str As String, nAscii As Integer
    For Each Cell In Selection.Cells
        If Cell.HasFormula = False Then
            str = Trim(CStr(Cell))
            If Len(str) > 0 Then
                nAscii = Asc(Left(str, 1))
                If nAscii < 33 Or nAscii = 160 Then
                    If Len(str) > 1 Then
                        str = Right(str, Len(str) - 1)
                    Else
                        str = ""
                    End If
                End If
            End If
            Cell = str
        End If
    Next
End Sub

' This macro function counts unique instances in selected column
Sub CountUniqueInSelectedColumn()
    Dim count As Integer
    Dim i, c, j As Integer
    Dim Rng As Range
    Set Rng = Selection
    c = 0
    count = 0
    For Each Cell In Selection.Cells
        For Each Rngcell In Rng
            If Rngcell.Value = Cell.Value Then
                MsgBox (Rngcell.Value)
            End If
        Next
    Next
End Sub

'This macro function attempts to read characters and formulate a potential password.
Sub DescryptProtectedSheet()
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
        MsgBox "One usable password is " & Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub

'This macro function converts URL text into hyperlinks
Public Sub ConvertToHyperlinks()
  Dim Cell As Range
  For Each Cell In Intersect(Selection, ActiveSheet.UsedRange)
    If Cell <> "" Then
      ActiveSheet.Hyperlinks.Add Cell, Cell.Value
    End If
  Next
End Sub
