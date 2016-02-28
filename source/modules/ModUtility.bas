Option Compare Database

Public Sub ErrorLog(strProc As String, ByVal lngErr As Long, ByVal strError As String)
'-----------------------------------------------------------
' Inputs: Name of the procedure trapping the error
'         Err value
'         Error string
' Output: Writes an entry to ErrorLog table
' Created By: JLV 01/31/95
' Last Revised: JLV 01/31/95
'-----------------------------------------------------------
Dim Db As Database
Dim rstE As Recordset
Dim strFrmName As String
Dim strCtlName As String
Dim lngErrSav As Long
Dim strErrSav As String
  
    On Error Resume Next
    lngErrSav = lngErr
    strErrSav = strError
    
    Set Db = CurrentDb()
    Set rstE = Db.OpenRecordset("ErrorLog")
    
    rstE.AddNew
    strFrmName = Screen.ActiveForm.name
    rstE!CurrentForm = strFrmName
    ' Added trap to avoid Fault when screen is blank
    If Not IsNothing(strFrmName) Then
        strCtlName = Screen.ActiveControl.name
    End If
    rstE!CurrentControl = strCtlName
    rstE!ActiveForms = Forms.Count
    rstE!UserName = CurrentUser
    rstE!Date = Now
    rstE!CallingProcedure = strProc
    rstE!ErrorCode = lngErrSav
    rstE!ErrorText = strErrSav
    rstE.Update
    rstE.Close

End Sub

Public Function IsFormLoaded(ByVal strFormName As String) As Integer
'-----------------------------------------------------------
' Inputs: Name of the form to test
' Outputs: True = form is in Forms collection; False = it ain't
' Created By: JLV 01/31/95
' Last Revised: JLV 01/31/95
'-----------------------------------------------------------

    On Error GoTo IsFormLoaded_Err

    IsFormLoaded = (SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0)

IsFormLoaded_Exit:
    On Error GoTo 0
    Exit Function

IsFormLoaded_Err:
    IsFormLoaded = False
    Err.Clear
    Resume IsFormLoaded_Exit

End Function

Public Function IsNothing(ByVal varValueToTest) As Integer
'-----------------------------------------------------------
' Does a "nothing" test based on data type.
'   Null = nothing
'   Empty = nothing
'   Number = 0 is nothing
'   String = "" is nothing
'   Date/Time is never nothing
' Inputs: A value to test for logical "nothing"
' Outputs: True = value passed is a logical "nothing", False = it ain't
' Created By: JLV 01/31/95
' Last Revised: JLV 01/31/95
'-----------------------------------------------------------
Dim intSuccess As Integer

    On Error GoTo IsNothing_Err
    IsNothing = True

    Select Case VarType(varValueToTest)
        Case 0      ' Empty
            GoTo IsNothing_Exit
        Case 1      ' Null
            GoTo IsNothing_Exit
        Case 2, 3, 4, 5, 6  ' Integer, Long, Single, Double, Currency
            If varValueToTest <> 0 Then IsNothing = False
        Case 7      ' Date / Time
            IsNothing = False
        Case 8      ' String
            If (Len(varValueToTest) <> 0 And varValueToTest <> " ") Then IsNothing = False
    End Select


IsNothing_Exit:
    On Error GoTo 0
    Exit Function

IsNothing_Err:
    IsNothing = True
    Resume IsNothing_Exit

End Function
Public Function SvuotaTab(nomeTab)
On Error GoTo GestoreErrori
   DoCmd.SetWarnings False
   DoCmd.RunSQL "DELETE * FROM [" & nomeTab & "];"
   DoCmd.SetWarnings True
Exit Function
GestoreErrori:
CError.WriteErrorToTable
End Function
Public Function ImportaFile(PercorsoFile, TabellaDestinazione, HasHeader)
On Error GoTo GestoreErrori
SvuotaTab TabellaDestinazione
'///////////////////////////////////////
    If HasHeader = False Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, TabellaDestinazione, PercorsoFile, False
    Else
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, TabellaDestinazione, PercorsoFile, True
        ' DoCmd.TransferSpreadsheet , acSpreadsheetTypeExcel9, "" & LeftvFileName & "", vFileName, True
    End If
' DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "tblPackingListImport", Indirizzo, False, "E6:R500"

Exit Function

Exit Function
GestoreErrori:
' CError.WriteErrorToTable
End Function
Public Function con(oldfield As String) As String
' Toglie apice e *
nfld1 = Replace(oldfield, "'", " ")
nfld2 = Replace(nfld1, "*", " ")
' nfld3 = Replace(nfld2, "$", ",")

'myarray = Split(nfld3, ",")
'con = myarray(0) & "," & myarray(1) & "," & myarray(2)
con = nfld2
End Function
Function CountOfActivePart() As Long

Dim strSQL As String, rs As Recordset

    On Error GoTo tagError

    strSQL = "SELECT Count(ID_Articoli) AS CountOfActivePart " & _
        "FROM tblArticoli "
    
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.EOF Then
        CountOfActivePart = 0
    Else
        CountOfActivePart = rs!CountOfActivePart
    End If
    rs.Close: Set rs = Nothing

    On Error GoTo 0
    Exit Function

tagError:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CountOfActivePart of Module mdlService"
    
End Function

Function RecordCountWhere_TSB(strDatabase As String, strWhere As String) As Long
    On Error GoTo RecordCountWhere_TSB_Err
  ' Comments  : Calcola il numero di record di una tabella o query
  ' Parameters: strDatabase - path and name of database to look in or "" (blank string) for the current database
  '             strWhere - Where clause to count records in
  ' Returns   : number of records
  '
  Dim dbsTemp As Database
  Dim rstTmp As Recordset
  Dim strSQL As String
  
  If strDatabase = "" Then
    Set dbsTemp = CurrentDb()
  Else
    Set dbsTemp = DBEngine.Workspaces(0).OpenDatabase(strDatabase)
  End If

  strSQL = "SELECT COUNT(*) AS CountOfRecords FROM " & strWhere & ";"
  Set rstTmp = dbsTemp.OpenRecordset(strSQL)

  RecordCountWhere_TSB = rstTmp![CountOfRecords]

  rstTmp.Close
  dbsTemp.Close
  

Exit Function

RecordCountWhere_TSB_Err:
    MsgBox Err.Description
    Exit Function
    Resume
Exit Function

End Function



Function MinOfList(ParamArray varValues()) As Variant
' Min di una lista di valori
' Es.:
'    =MinOfList(5, -3, Null, 0, 2)
'
' oppure inserire in una nuova colonna Sql :
'    MaxOfList([OrderDate], [InvoiceDate], [DueDate])
    Dim i As Integer        'Loop controller.
    Dim varMin As Variant   'Smallest value found so far.

    varMin = Null           'Initialize to null

    For i = LBound(varValues) To UBound(varValues)
        If IsNumeric(varValues(i)) Or IsDate(varValues(i)) Then
            If varMin <= varValues(i) Then
                'do nothing
            Else
                varMin = varValues(i)
            End If
        End If
    Next

    MinOfList = varMin
End Function

Function MaxOfList(ParamArray varValues()) As Variant
' MAX di una lista di valori
' Es.:
'    =MaxOfList(5, -3, Null, 0, 2)
'
' oppure inserire in una nuova colonna Sql :
'    MaxOfList([OrderDate], [InvoiceDate], [DueDate])
    Dim i As Integer        'Loop controller.
    Dim varMax As Variant   'Largest value found so far.

    varMax = Null           'Initialize to null

    For i = LBound(varValues) To UBound(varValues)
        If IsNumeric(varValues(i)) Or IsDate(varValues(i)) Then
            If varMax >= varValues(i) Then
                'do nothing
            Else
                varMax = varValues(i)
            End If
        End If
    Next

    MaxOfList = varMax
End Function

Public Function fncSQLStr(varStr As Variant) As String
' 20160122 Marco
' Funzione che toglie l'apostrofo, altrimenti nelle INSERT Sql ho errore
' Es.
'
'strSQL = "INSERT INTO tbl" &
'    " (fld1, fld2)" & _
'    " VALUES ('" & fncSQLStr(str1) & "', '" & fncSQLStr(Me.tfFld2.Value) & "');"
If IsNull(varStr) Then
        fncSQLStr = ""
    Else
        fncSQLStr = Replace(Trim(varStr), "'", "''")
    End If

End Function
Public Function AutoBackup()
    Dim sDataFile As String, sDataFileTemp As String, sDataFileBackup As String
    Dim s1 As Long, s2 As Long
    'Dim fs As Object

    Dim dbCur As DAO.Database _
        , dbAnal As DAO.Database

    Dim mDbPath As String _
        , mDbName As String

    Set dbCur = DBEngine(0)(0)    'CurrentDB

    Set dbAnal = CurrentDb
    mDbPath = CurrentProject.Path
    mDbName = Dir(CurrentDb.name)
    sDataFile = CurrentProject.Path & mDbName

    sDataFileTemp = CurrentProject.Path & "\PortafoglioOrdini2-15-Temp.accdb"
    sDataFileBackup = CurrentProject.Path & "\PortafoglioOrdini2-15-BackUp-" & Format(Now, "YYYY-MM-DD HHMMSS") & ".accdb"

    DoCmd.Hourglass True

    'get file size before compact
    Open sDataFile For Binary As #1
    s1 = LOF(1)
    Close #1

    'backup data file
    '' ==========  Create the backup file
    '                            Set fs = CreateObject("Scripting.FileSystemObject")
    '                            MsgBox "Backup in progress. Please wait"
    '                            fs.CopyFile sDataFile, sDataFileBackup
    '                            Set fs = Nothing
    FileCopy sDataFile, sDataFileBackup

    'only proceed if data file exists
    If FileExists(sDataFileBackup) = True Then

        'compact data file to temp file
        On Error Resume Next
        Kill sDataFileTemp
        On Error GoTo 0
        DBEngine.CompactDatabase sDataFileBackup, sDataFileTemp

        If Dir(sDataFileTemp, vbNormal) <> "" Then
            'delete old data file data file
            Kill sDataFile

            'copy temp file to data file
            FileCopy sDataFileTemp, sDataFile

            'get file size after compact
            Open sDataFile For Binary As #1
            s2 = LOF(1)
            Close #1

            DoCmd.Hourglass False
            MsgBox "Compact complete " & vbCrLf & vbCrLf _
                   & "Size before: " & Round(s1 / 1024 / 1024, 2) & "Mb" & vbCrLf _
                   & "Size after:    " & Round(s2 / 1024 / 1024, 2) & "Mb", vbInformation
        Else
            DoCmd.Hourglass False
            MsgBox "ERROR: Unable to compact data file"
        End If

    Else
        DoCmd.Hourglass False
        MsgBox "ERROR: Unable to backup data file"
    End If

    DoCmd.Hourglass False
End Function
Public Function FaseEseguita(IDFase)
' es. FaseEseguita ("03")
    On Error GoTo GestoreErrori
    ModificaCampo "tblFasiElaboraDati", "eseguita", "NUM_FASE", True, IDFase
    Tempofase IDFase, False
    Exit Function
GestoreErrori:
    CError.WriteErrorToTable
End Function
Public Function ModificaCampo(NomeTabella, NomeCampo, CampoRicerca, ValoreCampo, ValoreRicerca)
On Error GoTo GestoreErrori
Dim Db As DAO.Database
Dim tabella As DAO.Recordset
Dim campo As DAO.Field
Set Db = CurrentDb
Set tabella = Db.OpenRecordset(NomeTabella, dbOpenDynaset)
Set campo = tabella.Fields(NomeCampo)
Do Until tabella.EOF
    CR = tabella.Fields(CampoRicerca)
    If CR = ValoreRicerca Then
        tabella.Edit
        campo = ValoreCampo
        tabella.Update
    End If
    tabella.MoveNext
Loop
tabella.Close
Db.Close
Exit Function
GestoreErrori:
CError.WriteErrorToTable
End Function
Public Function Tempofase(FASE, INIZIO As Boolean)
' es. Tempofase ("03", False)

On Error GoTo GestoreErrori
If INIZIO = True Then
    ModificaCampo "tblFasiElaboraDati", "INIZIO", "NUM_FASE", Now(), FASE
Else
    ModificaCampo "tblFasiElaboraDati", "FINE", "NUM_FASE", Now(), FASE
End If
Exit Function
GestoreErrori:
CError.WriteErrorToTable
End Function