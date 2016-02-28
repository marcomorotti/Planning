Option Compare Database

Public Function ClasseABC(NomeTabella, Chiave, CampoValore, CampoABC, CampoID)
'La funzione consente il calcolo della classe ABC di una serie di valori presenti in
'una tabella. I valori da passare alla funzione sono i seguenti (da fornire tra
'virgolette poichè si tratta di testo) :
'NomeTabella:Indicare il nome della tabella su cui si vuole calcolare la classe ABC,
'CampoValore: Indicare il nome del campo della tabella in cui è contenuto il valore espresso & _
               in percentuale su cui calcolare la classe ABC.
'CampoABC: Nome del campo della tabella in cui la funzione andrà a scrivere la classe ABC calcolata,
'Chiave: Nome del campo della tabella in base al quale calcolare la classe ABC, & _
        ad esempio se in una tabella abbiamo più codici per i quali vogliamo calcolare la classe ABC & _
        la chiave è costituita dal codice ovvero la funzione ricalcola la classe ABC per ogni & _
        codice presente nella tabella
' CampoID: indicare il campo con l'etichetta dei valori di cui si calcola la classe ABC.
'Consideriamo di avere una tabella in cui è indicato il fatturato medio mensile degli
'articoli venduti per poter utilizzare la funzione la tabella dovrà avere la seguente
'struttura MESE,ARTICOLO,FATTURATO,CLASSE ABC, la funzione va usata in questo modo:
' ClasseABC "FATTURATO MENSILE","MESE","FATTURATO","CLASSE ABC","ARTICOLO"
' ClasseABC("tblArticoli", "Num_Eventi_12", "SConsumo_12", "AbcConsumo", "Cod_Art")

Dim Db As DAO.Database
Dim tabella As DAO.Recordset
Dim campo As DAO.Field
Dim ClassiABC()
numrighe = DCount(Chiave, NomeTabella)
ReDim ClassiABC(numrighe, 5)
Set Db = CurrentDb
testosql = "SELECT [" & NomeTabella & "].[" & CampoID & "],[" & NomeTabella & "].[" & Chiave & "], " & _
           "[" & NomeTabella & "].[" & CampoValore & "] FROM [" & NomeTabella & "] " & _
           "ORDER BY [" & NomeTabella & "].[" & Chiave & "],[" & NomeTabella & "].[" & CampoValore & "] DESC;"
Set tabella = Db.OpenRecordset(testosql, dbOpenDynaset)
Do Until tabella.EOF
    t = t + 1
    ClassiABC(t, 0) = tabella.Fields(CampoID)
    ClassiABC(t, 1) = tabella.Fields(Chiave)
    ClassiABC(t, 2) = tabella.Fields(CampoValore)
    tabella.MoveNext
Loop
tabella.Close
Db.Close
For x = 1 To t
    filtro = "[" & Chiave & "]=" & ClassiABC(x, 1)
    totale = DSum(CampoValore, NomeTabella, filtro)
    ClassiABC(x, 3) = totale
    ClassiABC(x, 4) = ClassiABC(x, 2) / ClassiABC(x, 3)
Next x
For x = 1 To t
    If ClassiABC(x, 1) <> ClassiABC(x - 1, 1) Then
        cumulata = ClassiABC(x, 4)
        If cumulata <= 0.8 Then
            classe = "A"
        Else
            If cumulata <= 0.9 Then classe = "B" Else classe = "C"
        End If
    Else
        cumulata = cumulata + ClassiABC(x, 4)
            If cumulata <= 0.8 Then
                classe = "A"
            Else
                If cumulata <= 0.9 Then classe = "B" Else classe = "C"
            End If
    End If
    ClassiABC(x, 5) = classe
Next x
'Set db = CurrentDb
'Set tabella = db.OpenRecordset(NomeTabella, dbOpenDynaset)
'Set campo = tabella.Fields(CampoABC)
'Do Until tabella.EOF
'    ID = tabella.Fields(CampoID)
'    For X = 1 To t
'        If ID = ClassiABC(X, 0) Then VCAMPO = ClassiABC(X, 5)
'    Next X
'    tabella.Edit
'    campo = VCAMPO
'    tabella.Update
'    tabella.MoveNext
'Loop
'tabella.Close
'db.Close
End Function