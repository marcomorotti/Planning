Option Compare Database   'Utilizza il tipo di ordinamento del database per i confronti fra stringhe
Option Explicit


Function leggiChiave(schiave As String)
'Legge nella tabella dei parametri se esiste il record chiave,
'se sì ne restituisce il valore altrimenti null


    leggiChiave = DLookup("Valore", "tblParametri", "chiave =""" & schiave & """")
End Function

Function scrivichiave(schiave As String, svalore)
'scrive nella tabella dei parametri con la chiave passata il valore passato
Dim miaNota As Variant
Dim ssql As String
'Stop

On Error GoTo scrivichiave_error

'Verifico che chiave e valore siano validi
If IsNull(schiave) Or schiave = "" Or IsNull(svalore) Or svalore = "" Then
    scrivichiave = False
    Exit Function
End If


Dim miodb As Database
Set miodb = DBEngine.Workspaces(0).Databases(0)

'leggo un eventuale valore per la nota
miaNota = Nz(DLookup("Commento", "tblParametri", "CHIAVE=""" & schiave & """ "), "")

ssql = "Delete * from tblParametri where chiave =""" & schiave & """"
miodb.Execute ssql

ssql = "insert into tblParametri(chiave,valore,commento) values (""" & schiave & """ , """ & svalore & """,""" & miaNota & """ )"
miodb.Execute ssql

scrivichiave = True
Exit Function

scrivichiave_error:
    MsgBox Error$
    scrivichiave = False
    Exit Function

End Function


Function leggiChiaveLocale(schiave As String)
'Legge nella tabella dei parametri locali se esiste il record chiave,
'se sì ne restituisce il valore altrimenti null
    leggiChiaveLocale = DLookup("Valore", "tblParametriLocali", "chiave =""" & schiave & """")
End Function


Function scrivichiaveLocale(schiave As String, svalore)
'scrive nella tabella dei parametri locali con la chiave passata il valore passato
Dim miaNota As Variant
Dim ssql As String
'Stop

On Error GoTo scrivichiaveLocale_error

'Verifico che chiave e valore siano validi
If IsNull(schiave) Or schiave = "" Or IsNull(svalore) Or svalore = "" Then
    scrivichiaveLocale = False
    Exit Function
End If


Dim miodb As Database
Set miodb = DBEngine.Workspaces(0).Databases(0)

'leggo un eventuale valore per la nota
miaNota = Nz(DLookup("Commento", "tblParametriLocali", "CHIAVE=""" & schiave & """ "), "")

ssql = "Delete * from tblParametriLocali where chiave =""" & schiave & """"
miodb.Execute ssql

ssql = "insert into tblParametriLocali(chiave,valore,commento) values (""" & schiave & """ , """ & svalore & """,""" & miaNota & """ )"
miodb.Execute ssql

scrivichiaveLocale = True
Exit Function

scrivichiaveLocale_error:
    MsgBox Error$
    scrivichiaveLocale = False
    Exit Function

End Function