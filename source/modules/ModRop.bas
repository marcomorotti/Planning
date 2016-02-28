Option Compare Database

Public Function ROP(Lead_time, AvgConsumoMese) As Double
' Inserire Lead_Time[mese]
Dim RopResult As Double
Dim Lead_Time_Mese As Double
Lead_Time_Mese = Lead_time / 22

RopResult = (Lead_Time_Mese * AvgConsumoMese)

If RopResult > 0 And RopResult <= 1 Then
        ROP = 1
' 20160127 was:
'If RopResult > 0 And RopResult <= 2 Then
'        ROP = 2
    Else
        ROP = Round(RopResult, 0)
    End If
    ' ROP = 100 * (Round(100 * RopResult, 0) / 100)
End Function
' ROP = Round(RopResult, 0)
' usare InverseCDF(Ls/100)