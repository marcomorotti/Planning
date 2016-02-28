Option Compare Database

Public Function ROQ(MesiCopertura, AvgConsumoMese, Cs_Csc) As Double
    Dim db0 As Database
    Dim Costo_unit_ord As Double
    Dim RoqResult As Double
    Dim DatiGen As New ADODB.Recordset
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    If IsNull(MesiCopertura) Then
        RoqResult = 0
    ElseIf Mid(MesiCopertura, 1, 1) = "F" Then
        RoqResult = Val(Mid(MesiCopertura, 2))
    Else
        Select Case Val(MesiCopertura)
        Case Is = 0    'Formula Lotto Economico
        'memorizzo il Costo di Acquisto
            DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
            Costo_unit_ord = DatiGen.Fields("Costo_Unit_Ord")
            DatiGen.Close
            RoqResult = Sqr((2 * (AvgConsumoMese * 12) * Costo_unit_ord) / (Cs_Csc * 0.68 * 0.21))
        Case Is > 0    'Formula Copertura
            RoqResult = AvgConsumoMese * Val(MesiCopertura)
        End Select
    End If
    
    
    If RoqResult > 0 And RoqResult < 1 Then
        RoqResult = 1
    ' 20151207 Se sono coperto dal ROQ per 2 anni (T_Tra_Ordini > 100 settimane - 52 settimane in un anno)
    ElseIf (RoqResult / (AvgConsumoMese * 12)) * 52 > 100 Then
       RoqResult = AvgConsumoMese * 24
    End If
    
    ' Arrotonda a 0
    ROQ = Round(RoqResult, 0)
       
    
    
End Function
'Public Function ROQ(MesiCopertura, AvgConsumoMese, Cs_Csc) As Double
'Dim db0 As Database
'Dim Costo_unit_ord As Double
'Dim RoqResult As Double
'Dim DatiGen As New ADODB.Recordset
'Set db0 = CurrentDb
'Set conn = CurrentProject.Connection
'
''memorizzo il Costo di Acquisto
'DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
'    Costo_unit_ord = DatiGen.Fields("Costo_Unit_Ord")
'DatiGen.Close
'
'Select Case MesiCopertura
'    Case Is = 0 'Formula Lotto Economico
'        RoqResult = Sqr((2 * (AvgConsumoMese * 12) * Costo_unit_ord) / (Cs_Csc * 0.68 * 0.21))
'        Case Is > 0 'Formula Copertura
'        RoqResult = AvgConsumoMese * MesiCopertura
'    End Select
'    If RoqResult > 0 And RoqResult < 1 Then
'        ROQ = 1
'    Else
'        ROQ = Round(RoqResult, 0)
'    End If
'End Function