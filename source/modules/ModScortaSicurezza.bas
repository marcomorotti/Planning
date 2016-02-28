Option Compare Database
' ********** Code Start **************
'Questo codice è stato scritto da Marco Morotti
' Inserire LivelloServizio in % es. 80 e Lt in gg
Public Function ScortaSicurezza(LivelloServizio, Lead_time, AvgConsumoMese, _
    DevStdConsumoMese) As Double
' Inserire Lead_Time[mese]
Dim SSResult As Double
Dim Lead_Time_Mese As Double
Lead_Time_Mese = Lead_time / 22
If LivelloServizio < 50 Then
    ScortaSicurezza = 0
Else
    SSResult = InverseCDF(LivelloServizio / 100) * (Sqr(Lead_Time_Mese * _
            (DevStdConsumoMese) ^ 2))
    ScortaSicurezza = Round(SSResult, 0)
End If
End Function