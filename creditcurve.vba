Function INTSPOT(spots, year)

Dim i As Integer
Dim spotnum As Integer

spotnum = spots.Rows.Count

If Application.WorksheetFunction.Count(spots) = 1 Then
INTSPOT = spots
Else
If yaer <= spots(1, 1) Then
INTSPOT = spots(1, 2)

ElseIf year >= spot(spotnum, 1) Then
INTSPOT = spots(spotnum, 2)
Else
Do
i = i + 1
Loop Until spots(i, 1) > year
INTSPOT = spots(i - 1, 2) + (spots(i, 2) - spots(i - 1, 2)) * (year - spots(i - 1)) / (spots(i, 1) - spots(i - 1, 1))
End If
End If


End Function
