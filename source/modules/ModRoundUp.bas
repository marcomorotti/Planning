Option Compare Database

Public Function RoundUp(dblNumToRound As Double, lMultiple As Long) As Double
    Dim asDec   As Variant
    Dim rounded As Variant
    If lMultiple = 0 Then
        RoundUp = dblNumToRound
        Exit Function
    End If
    asDec = CDec(dblNumToRound) / lMultiple
    rounded = Int(asDec)

    If rounded <> asDec Then
       rounded = rounded + 1
    End If
    RoundUp = rounded * lMultiple
End Function