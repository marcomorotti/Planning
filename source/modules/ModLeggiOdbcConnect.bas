Option Compare Database   'Utilizza il tipo di ordinamento del database per i confronti fra stringhe
Option Explicit

Function LeggiOdbcConnect()
'Legge nella tabella login dbo_AdminTable se esiste il record chiave,
'se sì ne restituisce il valore altrimenti prende il default
    
    If IsNull(DLookup("OdbcConnect", "dbo_AdminTable", "LoginID =""" & GetUser() & """")) Then
        LeggiOdbcConnect = leggiChiave("OdbcConnect")
    Else
        LeggiOdbcConnect = DLookup("OdbcConnect", "dbo_AdminTable", "LoginID =""" & GetUser() & """")
    End If
End Function