Option Compare Database

Public Function CrossAnalysisNumRighe(NomeTabella, CampoCosti, ClasseGiacenza, ClasseConsumo)

'---------------------------------------------------------------------------------------
' Procedura : Matrice ABC
'             es. Call CrossAnalysisNumRighe("tblArticoli", "Cs_Csc", "A", "A")
' ' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2013-Aprile-11          inizio
'---------------------------------------------------------------------------------------
'Definizioni
'***********************************************************************
Dim Db As DAO.Database
Dim rst As DAO.Recordset
' VarWhere = "SELECT * FROM " & Nometabella & _
            " WHERE AbcGiacenza = '" & ClasseGiacenza & "' AND AbcConsumo = '" & ClasseConsumo & "';"
            
Set rst = DBEngine(0)(0).OpenRecordset("SELECT * FROM " & NomeTabella & _
            " WHERE AbcGiacenza = '" & ClasseGiacenza & "' AND AbcConsumo = '" & ClasseConsumo & "' AND Giac_Media <> 0;") 'Morotti 11/9/13
' 20130827
If (rst.BOF And rst.EOF) Then 'se non trova record esce
    CrossAnalysisNumRighe = 0
    Exit Function
End If
rst.MoveLast
numrighe = rst.RecordCount 'Totale codici
rst.Close
Set rst = Nothing
CrossAnalysisNumRighe = numrighe
End Function
Public Function CrossAnalysisTotGiacenza(NomeTabella, CampoCosti, ClasseGiacenza, ClasseConsumo)
'---------------------------------------------------------------------------------------
' Procedura : Matrice ABC
'             es. Call CrossAnalysisTotGiacenza("tblArticoli", "Cs_Csc", "A", "A")
' ' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2013-Aprile-11          inizio
'---------------------------------------------------------------------------------------
'Definizioni
'***********************************************************************
Dim Db As DAO.Database
Dim rst As DAO.Recordset
Set Db = CurrentDb
' TestoSql = "SELECT [" & NomeTabella & "].[Giac_Media] * [" & NomeTabella & "].[" & CampoCosti & "], " & _
           "[" & NomeTabella & "].[SConsumo_12] * [" & NomeTabella & "].[" & CampoCosti & "] FROM [" & NomeTabella & "] " & _
           " WHERE AbcGiacenza = '" & ClasseGiacenza & "' AND AbcConsumo = '" & ClasseConsumo & "';"

' Totale Giacenza
testosql = "SELECT Sum([" & NomeTabella & "].[Giac_Media] * [" & NomeTabella & "].[" & CampoCosti & "]) FROM [" & NomeTabella & "] " & _
           " WHERE AbcGiacenza = '" & ClasseGiacenza & "' AND AbcConsumo = '" & ClasseConsumo & "';"

Set rst = DBEngine(0)(0).OpenRecordset(testosql)
 TotaleGiacenza = rst(0)
rst.Close
Set rst = Nothing
CrossAnalysisTotGiacenza = TotaleGiacenza
End Function
Public Function CrossAnalysisTotConsumo(NomeTabella, CampoCosti, ClasseGiacenza, ClasseConsumo)
'---------------------------------------------------------------------------------------
' Procedura : Matrice ABC
'             es. Call CrossAnalysis("tblArticoli", "Cs_Csc", "A", "A")
' ' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2013-Aprile-11          inizio
'---------------------------------------------------------------------------------------
'Definizioni
'***********************************************************************
Dim Db As DAO.Database
Dim rst As DAO.Recordset
Set Db = CurrentDb
' Totale Consumo
' SOSTITUITO CAMPO SConsumo_12 CON SConsumo  22/05/2013
testosql = "SELECT Sum([" & NomeTabella & "].[SConsumo_12] * [" & NomeTabella & "].[" & CampoCosti & "]) FROM [" & NomeTabella & "] " & _
           " WHERE AbcGiacenza = '" & ClasseGiacenza & "' AND AbcConsumo = '" & ClasseConsumo & "';"

Set rst = DBEngine(0)(0).OpenRecordset(testosql)
 TotaleConsumo = rst(0)
rst.Close
Set rst = Nothing
Db.Close

CrossAnalysisTotConsumo = TotaleConsumo
End Function