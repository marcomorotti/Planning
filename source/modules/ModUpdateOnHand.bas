Option Compare Database

Public Function UpdateOnHand()
Dim ws As Workspace
    Dim Db As Database
    Dim db0 As Database    ' ******* DATA BASE CORRENTE ***********
    Dim Lconnect As String
    Dim MyQuery As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim Response As Integer
    Dim conn As ADODB.Connection
    Dim intI As Double

    ' *** CONNESSIONE a Oracle

    ' On Error GoTo Err_Execute

    'Use {Microsoft ODBC for Oracle} ODBC connection
    'Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
    Lconnect = LeggiOdbcConnect
    'Point to the current workspace
    Set ws = DBEngine.Workspaces(0)

    'Connect to Oracle
    Set Db = ws.OpenDatabase("", False, True, Lconnect)
    ' Setto il tempo di QueryTimeOut a 240 min
    Db.QueryTimeout = 240


    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    ' GoTo Start
    ' *** Inizio caricamento GIACENZA.sql
    '
    'Visualizza lo Status Meter
    Call acbInitMeter("1-Aggiorna GIACENZA da ORACLE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella dati
    conn.Execute "Delete * From tblInventoryControl"

    MyQuery = "SELECT CD_ART, QTY_SALE, QTY_STOCK, QTY_ACQ " & _
              " FROM on_hand; "


    Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    If Not rst.BOF Then    'se ci sono record nel recordset
        rst.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = rst.RecordCount
        rst.MoveFirst
    End If
    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qryInventoryControlIsrt")
        qm!iCD_ART = rst![Cd_Art]
        If IsNull(rst![QTY_SALE]) Then
            QTY_SALE = 0
        Else
            QTY_SALE = rst![QTY_SALE]
        End If
        qm!iQTY_SALE = QTY_SALE
        If IsNull(rst![QTY_STOCK]) Then
            QTY_STOCK = 0
        Else
            QTY_STOCK = rst![QTY_STOCK]
        End If
        qm!iQTY_STOCK = QTY_STOCK
        If IsNull(rst![QTY_ACQ]) Then
            QTY_ACQ = 0
        Else
            QTY_ACQ = rst![QTY_ACQ]
        End If
        qm!iQTY_ACQ = QTY_ACQ
        qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
        ' Debug.Print ("Costo: " & rst![Cs_Csc])
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    rst.Close

    'Close lo Status Meter
    Call acbCloseMeter
End Function