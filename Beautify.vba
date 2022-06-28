Sub Beautify()

Dim R As Range, Idx As Long

    Set R = Selection
    Idx = 1

    Do While Idx < R.Rows.Count                           ' count dynamically changes as we delete rows
        If R(Idx + 1, 1) = "" Then                      ' found a break line looking 1 down
            R(Idx, 2) = R(Idx, 2) & " " & R(Idx + 1, 2)   ' append to current
            R(Idx + 1, 1).EntireRow.Delete                ' delete following but do not count up Idx
        Else
            Idx = Idx + 1                                 ' this one is clean, advance
        End If
    Loop
End Sub

'r(idx,2) means the position of the index and the column number where you want the data to be pasted
