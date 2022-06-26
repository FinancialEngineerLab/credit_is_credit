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


Function MYPRICE(settlemnt As Date, maturity As Date, rate, spots, notional, freq As Integer, _
Optional compound As Integer, Optional fromdate As Date, Optional basis As Integer)


Dim t As Date
Dim y As Double

If compound = 0 Then
compound = freq

If fromdate = 0 Then
fromdate = settlement
If fromdate > maturity Or settlement > maturity Then
End


t = maturity
y = YearFrac(settlement, maturity, basis)

MYPRICE = (notional + notional * rate / freq) / (1 + INTSPOT(spots, y) / compound) ^ (y * compound)

t = CoupPcd(t - 1, maturity, freq, basis)
Do While t > settlement And t >= fromdate
y = YearFrac(settlement, t, basis)
MYPRICE = MYPRICE + rate / freq * notional / (1 + INTSPOT(spots, y) / compound) ^ (y * compound)
t = CoupPcd(t - 1, maturity, freq, basis)
Loop




End Function




Function ACI(settlement As Date, maturity As Date, rate, freq As Integer, Optional basis As Integer)

If settlement < maturity Then
ACI = 100 * rate / freq * (1 - CoupDaysNc(settlement, maturity, freq, basis) / CoupDays(settlement, maturity, freq, basis))
End If

If ACI = 0 Or settlement = maturity Then
ACI = 100 * rate / freq

End Function
