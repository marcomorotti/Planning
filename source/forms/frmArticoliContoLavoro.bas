Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =5272
    ItemSuffix =23
    Left =6870
    Top =3060
    Right =12135
    Bottom =4935
    RecSrcDt = Begin
        0xa3aa30f84e40e140
    End
    Caption ="Articoli Obsoleti"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            TextFontFamily =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            SpecialEffect =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin FormHeader
            Height =577
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Top =45
                    Width =3615
                    Height =405
                    FontSize =14
                    ForeColor =255
                    Name ="Text12"
                    Caption ="Art. CONTO LAVORO"
                    LayoutCachedTop =45
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =450
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4639
                    Width =576
                    Height =576
                    FontSize =8
                    FontWeight =400
                    Name ="cmdExit"
                    Caption ="(demo)"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Tag ="ShiftLeftNewUnit,ShiftLeftPermLocn"
                    ControlTipText ="Esce..."
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4639
                    LayoutCachedWidth =5215
                    LayoutCachedHeight =576
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    TextFontFamily =34
                    Left =4062
                    Width =577
                    Height =577
                    TabIndex =1
                    Name ="cmdExcel"
                    Caption ="E&xcel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Export dati ..."
                    UnicodeAccessKey =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000030000000b00000014000000190000001d ,
                        0x000000230000002a000000320000003800000040000000470000004c0000003b ,
                        0x0000002700000005000000000000001000000020000000400000005000000050 ,
                        0x4040407060686080a0a0a0a0b0b0b0b0c0c8c0c0d0e0e0e0e0f0e0f0d0d8d0c0 ,
                        0x0000004000000020106010c0105810d080a080f0c0d0c0f0d0e0d0ffffffffff ,
                        0xfffffffffffffffff0fffffff0fffffff0fffffff0fffffff0fff0ffc0d8c0f0 ,
                        0x0000004000000020106010f0a0b8a0ffc0d0c0ffffffffffffffffffffffffff ,
                        0xfffffffff0fffffff0ffffff90c0a0ff107030ff107030fff0f8f0ffe0e8e0ff ,
                        0x0000004000000020106810f0b0d0c0ff80a080ffffffffffc0e0c0ff80c080ff ,
                        0x70b870ff70b870ff60b070ff107030ff107030ff60a870ffe0f0e0fff0f8f0ff ,
                        0x0000004000000020106810f0c0d0c0ff407840ffffffffffffffffffc0e0c0ff ,
                        0x70b870ff70b870ff208040ff107030ff60a070ff60b070ffa0d0a0fff0ffffff ,
                        0x4040406000000020107010f0c0d0c0ff206020fff0f0f0ffffffffffffffffff ,
                        0xe0f0e0ff409860ff107030ff207840ffb0e0c0fffffffffff0fffffff0ffffff ,
                        0x7070707000000020207020f0c0d0c0ff206020ffb0c8b0ffffffffffffffffff ,
                        0xc0d8d0ff107030ff107030ff60a860ff70b870ffa0d0a0fff0fff0fff0ffffff ,
                        0x9098909000000020207820f0c0d0c0ff206820ff90a890fffffffffff0f8f0ff ,
                        0x107030ff107030ff90c0a0ffe0f0e0ff70b870ff70b870ff90c890fff0f8f0ff ,
                        0xb0b0b0a000000020207820f0d0e0d0ff70a870ff90b090ffffffffff409060ff ,
                        0x308050ff60a070ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xd0d0d0a000000020208020f0d0e8d0ff80b880ff80b080ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xf0f0f0b000000020208020e0d0e8d0ffa0c8a0ff80b880ffd0e8d0ffd0e0d0ff ,
                        0xd0e0d0ffd0e0d0ffc0d8c0ffb0d8b0ffa0d0a0ff90c8a0ffd0e8d0ff50a050e0 ,
                        0xffffff2000000000208020a0a0c890ffe0f0e0ff90c890ff90c890ff90c890ff ,
                        0x90c890ff90c890ff90c890ff90c890ff90c890ff90c890fff0f8f0ff208020d0 ,
                        0x000000000000000030882030308820ffc0e0c0fff0f8f0ffb0d8b0ffa0d0a0ff ,
                        0xa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ff90d0a0fff0f8f0ff308820d0 ,
                        0x00000000000000000000000030882060308820f090c890ffd0e8d0fff0f8f0ff ,
                        0xf0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0ff308820d0 ,
                        0x0000000000000000000000000000000030902020309020a0309020e0309020ff ,
                        0x309020ff309020ff309020ff309020ff309020ff309020ff309020ff309020d0 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4062
                    LayoutCachedWidth =4639
                    LayoutCachedHeight =577
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =3486
                    Width =576
                    Height =576
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="cmdApriDir"
                    Caption ="(demo)"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="ShiftLeftNewUnit,ShiftLeftPermLocn"
                    ControlTipText ="Apre direttorio export ..."
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000708890ff608090ff607880ff507080ff506070ff405860ff ,
                        0x404850ff303840ff203030ff202030ff101820ff101010ff101020ff00000000 ,
                        0x0000000000000000708890ff90a0b0ff70b0d0ff0090d0ff0090d0ff0090d0ff ,
                        0x0090c0ff1088c0ff1080b0ff1080b0ff2078a0ff207090ff204860ff20303050 ,
                        0x0000000000000000808890ff80c0d0ff90a8b0ff80e0ffff60d0ffff50c8ffff ,
                        0x50c8ffff40c0f0ff30b0f0ff30a8f0ff20a0e0ff1090d0ff206880ff202830b0 ,
                        0x00000000000000008090a0ff80d0f0ff90a8b0ff90c0d0ff70d8ffff60d0ffff ,
                        0x60d0ffff50c8ffff50c0ffff40b8f0ff30b0f0ff30a8f0ff1088d0ff204860ff ,
                        0x10283020000000008090a0ff80d8f0ff80c8e0ff90a8b0ff80e0ffff70d0ffff ,
                        0x60d8ffff60d0ffff60d0ffff50c8ffff40c0f0ff40b8f0ff30b0f0ff206880ff ,
                        0x10486090000000008098a0ff90e0f0ff90e0ffff90a8b0ff90b8c0ff70d8ffff ,
                        0x60d8ffff60d8ffff60d8ffff60d0ffff50d0ffff50c8ffff40b8f0ff30a0e0ff ,
                        0x406070f0506070308098a0ff90e0f0ffa0e8ffff80c8e0ff90a8b0ff80e0ffff ,
                        0x80e0ffff80e0ffff80e0ffff80e0ffff80e0ffff80e0ffff70d8ffff70d8ffff ,
                        0x50a8d0ff506070a090a0a0ffa0e8f0ffa0e8ffffa0e8ffff90b0c0ff90b0c0ff ,
                        0x90a8b0ff90a8b0ff80a0b0ff80a0b0ff8098a0ff8098a0ff8090a0ff8090a0ff ,
                        0x808890ff708890ff90a0b0ffa0e8f0ffa0f0ffffa0e8ffffa0e8ffff80d8ffff ,
                        0x60d8ffff60d8ffff60d8ffff60d8ffff60d8ffff60d8ffff708890ff00000000 ,
                        0x000000000000000090a0b0ffa0f0f0ffb0f0f0ffa0f0ffffa0e8ffffa0e8ffff ,
                        0x70d8ffff90a0a0ff8098a0ff8098a0ff8090a0ff809090ff708890ff00000000 ,
                        0x000000000000000090a8b0ffa0d0e0ffb0f0f0ffb0f0f0ffa0f0ffffa0e8ffff ,
                        0x90a0b0ff80a8b0800000000000000000000000000000000000000000906850ff ,
                        0x906850ff906850ff90a8b05090a8b0ff90a8b0ff90a8b0ff90a8b0ff90a8b0ff ,
                        0x90a8b090000000000000000000000000000000000000000000000000a0787050 ,
                        0x906850ff906850ff000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000907860ff9068506000000000a0787010a09080ff ,
                        0xa0887050907860ff000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000a0988040a09080ffa08880ffb09880ffa0908080 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3486
                    LayoutCachedWidth =4062
                    LayoutCachedHeight =576
                End
            End
        End
        Begin Section
            Height =1322
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1485
                    Top =330
                    Width =2505
                    Height =765
                    FontSize =12
                    ForeColor =255
                    Name ="Etichetta22"
                    Caption ="Premi tasto Excel per esportare i dati"
                    LayoutCachedLeft =1485
                    LayoutCachedTop =330
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1095
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit


Private Sub cmdApriDir_Click()
    Call fHandleFile(CurrentProject.Path & "\Export", WIN_NORMAL)
End Sub

Private Sub cmdExcel_Click()
Dim Response As Integer
  
' Dim nRecords As Integer 'Long
On Error GoTo Err_cmdExcel_Click:

Response = MsgBox("Vuoi esportare in Excel gli ARTICOLI di Conto Lavoro? ", _
    vbYesNo, "Continue")
If Response = vbYes Then
    Response = MsgBox("Vuoi importare il CONTO LAVORO da Oracle? ", vbYesNo, "Continue")
    If Response = vbYes Then
        ' ==== Importa via ODBC i dati da Oracle  ========
        Dim ws As Workspace
        Dim Db As Database
        Dim db0 As Database ' ******* DATA BASE CORRENTE ***********
        Dim Lconnect As String
        Dim MyQuery As String
        Dim qm As QueryDef
        Dim rst As Object
        Dim rs As Recordset
        Dim conn As ADODB.Connection
        Dim intI As Double
        ' Verificare i Dim
        Dim FromDate As Variant, ToDate As Variant
        Dim StartDate As Variant, EndDate As Variant
        Dim DatiGen As New ADODB.Recordset
        ' Dim CatEventi As New adodb.Recordset
            ' Dim CalMag As New adodb.Recordset
            ' Dim CalGiac As New adodb.Recordset
            Dim Art As DAO.Recordset
            Dim Cmd As New ADODB.Command
            Dim bqry As DAO.QueryDef
            Dim brs As DAO.Recordset
        
        ' Dim nRecords As Integer 'Long
        On Error GoTo Err_cmdImportDataOracle_Click:
        
        
        Set db0 = CurrentDb
        Set conn = CurrentProject.Connection
        
        ' *** CONNESSIONE a Oracle
        
        ' On Error GoTo Err_Execute
        
        'Use {Microsoft ODBC for Oracle} ODBC connection
            'Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
            Lconnect = LeggiOdbcConnect
            'Point to the current workspace
            Set ws = DBEngine.Workspaces(0)
            
            'Connect to Oracle
            Set Db = ws.OpenDatabase("", False, True, Lconnect)
            ' Setto il tempo di QueryTimeOut a 120 min
            Db.QueryTimeout = 240
        
        ' ****************************************************
        ' *** Inizio caricamento CONTO LAVORO ***
        ' ****************************************************
        
            'Visualizza lo Status Meter
            Call acbInitMeter("IN CONTO LAVORO da ORACLE", True)
            'Reset lo Status Meter
            intI = 0
            ' Cancella dati
            conn.Execute "Delete * From tblImportContoLavoro"
            
              MyQuery = "SELECT I.CD_ART, " & _
                        "A.DS_TEC, " & _
                        "I.tb_ubic, " & _
                        "I.qt_giac, " & _
                        "I.qt_imp " & _
                        "FROM grp.mag_articolo I, GRP.AN_ARTICOLO_GRP A " & _
                        "WHERE i.CD_ART = A.CD_ART " & _
                        "AND I.qt_imp > 0 " & _
                        "AND I.cd_soc = ' SP' " & _
                        "AND I.tb_ubic LIKE ('%*%')"
        
                        
                        
            Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
            Do While Not rst.EOF
               Set qm = db0.QueryDefs("qryContoLavoroIsrt")
               qm!iCD_ART = rst![Cd_Art]
               qm!iDescr_Art = rst![DS_TEC]
               qm!itb_ubic = rst![tb_ubic]
               qm!iQT_GIAC = rst![qt_giac]
               qm!iqt_imp = rst![qt_imp]
               qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
               qm.Execute
               rst.MoveNext
        
            ' Aggiorna lo Status Meter
               intI = intI + 1
                Call acbUpdateMeter(Int(intI))
            ' Attende un secondo
            DoEvents
            Loop
            rst.Close
            'Close lo Status Meter
            Call acbCloseMeter
            'db.Close
        ' *******************************************************************
        ' *** Inizio caricamento GIACENZA SP PER ARTICOLI IN CONTO LAVORO ***
        ' *******************************************************************
        
            'Visualizza lo Status Meter
            Call acbInitMeter("GIAC SP x ART IN CONTO LAVORO da ORACLE", True)
            'Reset lo Status Meter
            intI = 0
            ' Cancella dati
            conn.Execute "Delete * From tblImportContoLavoroGiacSp"
            
              MyQuery = "SELECT giacenza.CD_ART, " & _
                        "SUM (giacenza.qt_giac) AS QT_GIAC " & _
                        "FROM grp.mag_articolo giacenza " & _
                        "WHERE giacenza.qt_giac > 0 " & _
                        "AND giacenza.cd_soc = ' SP' " & _
                        "AND giacenza.tb_ubic IN ('SATT', 'SAUT', 'SDIR', 'XFGB', 'XFPL', ' RIC', 'XRESO', 'XROTRIC') " & _
                        "AND exists (select * " & _
                        "from grp.mag_articolo I " & _
                        "where i.cd_art = giacenza.cd_art " & _
                        "AND I.qt_imp > 0 " & _
                        "AND I.cd_soc = ' SP' " & _
                        "AND I.tb_ubic LIKE ('%*%')) " & _
                        "group by cd_art"
                        
                        
            Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
            Do While Not rst.EOF
               Set qm = db0.QueryDefs("qryContoLavoroGiacSpIsrt")
               qm!iCD_ART = rst![Cd_Art]
               qm!iQT_GIAC = rst![qt_giac]
               qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
               qm.Execute
               rst.MoveNext
        
            ' Aggiorna lo Status Meter
               intI = intI + 1
                Call acbUpdateMeter(Int(intI))
            ' Attende un secondo
            DoEvents
            Loop
            rst.Close
            'Close lo Status Meter
            Call acbCloseMeter
            'db.Close
        
Else
End If
    Dim strDataLine As String
    Dim intFile As Integer
    Dim filenm As String
    Dim i As Integer
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    
    ' Crea File di output
    intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
    "ArticoliContoLavoro_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "ArticoliContoLavoro_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "ArticoliContoLavoro_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "ArticoliContoLavoro_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Codice" & Chr(9) & "Descrizione" & Chr(9) & "Ubicazione" & _
        Chr(9) & "Giacenza Sp" & Chr(9) & "Giacenza CL" & Chr(9) & "Impegnato"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
        
    
    Set bqry = db0.QueryDefs("qryArticoliContoLavoro")
    Set brs = bqry.OpenRecordset
    
    Do While Not brs.EOF
    strDataLine = brs.Fields("Cd_art").Value & Chr(9) & _
        brs.Fields("Descr_art").Value & Chr(9) & brs.Fields("tb_ubic").Value & _
        Chr(9) & brs.Fields("GiacSp").Value & _
        Chr(9) & brs.Fields("qt_giac").Value & Chr(9) & brs.Fields("qt_imp").Value
    
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Loop
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "ArticoliContoLavoro_" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set db0 = Nothing
    Close #intFile
End If
Exit_cmdExcel_Click:
    Exit Sub

Err_cmdExcel_Click:
    MsgBox Err.Description
    Resume Exit_cmdExcel_Click
    
Exit_cmdImportDataOracle_Click:
    Exit Sub

Err_cmdImportDataOracle_Click:
    MsgBox Err.Description
    
    Resume Exit_cmdImportDataOracle_Click

    
End Sub


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdHelpReportCost_Click()
DoCmd.OpenForm "z Help Text for User", , , "zhID = 90"
End Sub
