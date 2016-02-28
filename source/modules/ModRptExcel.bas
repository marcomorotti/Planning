Option Compare Database

Public Sub GenerateExcelOrders(Varwhere As Variant, PrintOut As Boolean)
' // Per usare questa funzione aggiungere librerie Excel e Word
Dim xlApp As Excel.Application
Dim xlWb As Excel.Workbook
Dim xlWs As Excel.Worksheet
Dim xlrng As Excel.Range
Dim filenm As String
Dim invoiceamt As Currency

filenm = CurrentProject.Path & "\Template\OrdiniTemplate.xls"

Dim Db As DAO.Database
Dim cqry As DAO.QueryDef
Dim crs As DAO.Recordset
Dim bqry As DAO.QueryDef
Dim brs As DAO.Recordset
Dim brs1 As DAO.Recordset
Dim uinv As DAO.QueryDef
Dim pqry As DAO.QueryDef
Dim x As Long
Dim NumeratoreRmc As String
Dim i As Integer

Set xlApp = New Excel.Application
xlApp.Visible = True
Set xlWb = xlApp.Workbooks.Open(filenm)
Set xlWs = xlWb.Sheets("Data")
xlWs.name = "Data"

Set Db = CurrentDb
If IsNothing(Varwhere) Then
    Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryCOrdersExe")
Else
Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryCOrdersExe WHERE " & _
    Varwhere)
End If

' X = numero della riga da dove stampare
x = 6
If Not brs.EOF And Not brs.BOF Then
brs.MoveFirst
End If
While Not brs.EOF
xlWs.Cells(x, 2).Value = brs.Fields("Liv_Urgenza").Value
xlWs.Cells(x, 3).Value = brs.Fields("Numero_doc").Value
xlWs.Cells(x, 4).Value = brs.Fields("Cod_Cli").Value
xlWs.Cells(x, 5).Value = brs.Fields("Ds_Rag_soc").Value
xlWs.Cells(x, 6).Value = brs.Fields("Cod_Art").Value
xlWs.Cells(x, 7).Value = brs.Fields("Descrizione").Value
xlWs.Cells(x, 8).Value = brs.Fields("Qta_Ord_umv").Value
xlWs.Cells(x, 9).Value = brs.Fields("Data_Prev_Cons").Value
xlWs.Cells(x, 10).Value = brs.Fields("Qta_Cons_umv").Value
xlWs.Cells(x, 11).Value = brs.Fields("DispSp").Value
xlWs.Cells(x, 12).Value = brs.Fields("Impegnato").Value
xlWs.Cells(x, 13).Value = brs.Fields("DispAh").Value
xlWs.Cells(x, 14).Value = brs.Fields("Ord_Acq").Value
xlWs.Cells(x, 15).Value = brs.Fields("Cod_Forn").Value
xlWs.Cells(x, 16).Value = brs.Fields("Rag_Soc_Forn").Value
xlWs.Cells(x, 17).Value = brs.Fields("Data_Ordine").Value
xlWs.Cells(x, 18).Value = brs.Fields("Qta_Residua").Value
xlWs.Cells(x, 19).Value = brs.Fields("Data_Ric").Value
x = x + 1
brs.MoveNext
Wend
brs.Close
Set brs = Nothing
Set Db = Nothing
i = 1
' DoCmd.SetWarnings = False
Excel.Application.DisplayAlerts = False

If FileExists(TrailingSlash(CurrentProject.Path) & "\" & "Documents\" & Format(Date, "yyyymmdd") _
    & "Portafoglio_" & i & ".xls") = False Then
xlWb.SaveAs CurrentProject.Path & "\" & "Documents\" & Format(Date, "yyyymmdd") _
    & "Portafoglio_" & i & ".xls", FileFormat:=xlNormal, ConflictResolution:=xlOtherSessionChanges
Else
For i = 2 To 100
    If FileExists(TrailingSlash(CurrentProject.Path) & "\" & "Documents\" & Format(Date, "yyyymmdd") _
    & "Portafoglio_" & i & ".xls") = True Then
        GoTo NextOne
    End If
xlWb.SaveAs CurrentProject.Path & "\" & "Documents\" & Format(Date, "yyyymmdd") _
        & "Portafoglio_" & i & ".xls", ConflictResolution:=xlOtherSessionChanges, FileFormat:=xlNormal
'DoCmd.SetWarnings = False
MsgBox ("Data has been saved as Excel spreadsheet" & CurrentProject.Path & "\" _
    & "Documents\" & Format(Date, "yyyymmdd") & "Portafoglio_" & i & ".xls")
If PrintOut = True Then xlWs.PrintOut
' // Chiude Foglio Excel.
xlWb.Close
xlApp.Quit
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
Set Db = Nothing
Exit Sub
NextOne:
    Next i
End If
End Sub