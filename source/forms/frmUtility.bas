Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5340
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =9360
    Top =2790
    Right =14700
    Bottom =6285
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3955d11af5b9e340
    End
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Section
            Height =3510
            BackColor =13229799
            Name ="Corpo"
            AlternateBackColor =13229799
            Begin
                Begin Label
                    SpecialEffect =4
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =-30
                    Width =4635
                    Height =555
                    FontSize =20
                    FontWeight =700
                    BackColor =12311007
                    ForeColor =1845071
                    Name ="Etichetta7"
                    Caption ="UTILITY"
                    FontName ="Verdana"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =90
                    Top =645
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    ForeColor =1845071
                    Name ="cmdNewAssegnatario"
                    Caption ="DATI GENERALI"
                    OnClick ="=ApriMaschere(\"frmDatiGenerali\")"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2777
                    Top =645
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =1845071
                    Name ="CmdCatEventi"
                    Caption =" MOVIMENTAZIONE"
                    OnClick ="=ApriMaschere(\"frmNewCategoriaEventi\")"
                    FontName ="Gill Sans MT"

                    LayoutCachedLeft =2777
                    LayoutCachedTop =645
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1303
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =1845071
                    Name ="cmdNewTipoVeicolo"
                    Caption ="CATEG. COSTI"
                    OnClick ="=ApriMaschere(\"frmNewCategoriaCosti\")"
                    FontName ="Gill Sans MT"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1303
                    LayoutCachedWidth =2388
                    LayoutCachedHeight =1738
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4755
                    Top =15
                    Width =576
                    Height =576
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdClose"
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
                    ControlTipText ="Close the form"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2777
                    Top =1303
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =1845071
                    Name ="cmdTestOdbc"
                    Caption ="TEST ODBC"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2777
                    LayoutCachedTop =1303
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =1738
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1858
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =1845071
                    Name ="cmdLingue"
                    Caption ="LINGUE"
                    OnClick ="=ApriMaschere(\"frmLanguageMaintenance\")"
                    FontName ="Gill Sans MT"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =1858
                    LayoutCachedWidth =2388
                    LayoutCachedHeight =2293
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2834
                    Top =1870
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    ForeColor =1845071
                    Name ="cmdAggDescrArt"
                    Caption ="AGG. ART-DESCR-EN"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2834
                    LayoutCachedTop =1870
                    LayoutCachedWidth =5102
                    LayoutCachedHeight =2305
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =113
                    Top =2381
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    ForeColor =1845071
                    Name ="Comando14"
                    Caption ="OBSOLETI IMPORT"
                    OnClick ="=ApriMaschere(\"frmImport\")"
                    FontName ="Gill Sans MT"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =113
                    LayoutCachedTop =2381
                    LayoutCachedWidth =2381
                    LayoutCachedHeight =2816
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAggDescrArt_Click()
Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
' Dim rst As DAO.Recordset
'    Set dbs = CurrentDb ' Imposta il database corrente come attivo
'    ' Crea un oggetto TableDef nel database corrente.
'    Set tdf = dbs.CreateTableDef("tblArticoliDescrEstesa")
'    ' Imposta le proprietà Connect e SourceTableName per la tabella;
'    ' si tratta di un foglio di lavoro Excel.
' tdf.Connect = "Excel 8.0;DATABASE=" & CurrentProject.Path & "\DATA\tblArticoliDescrEstesa.xlsx;HDR=Yes"
'
' tdf.SourceTableName = "SqlResults$"
'
'' Accoda l'oggetto Tabledef all'insieme Tabledefs del database.
'    dbs.TableDefs.Append tdf
DoCmd.SetWarnings False
strSQL = "UPDATE tblArticoli INNER JOIN tblArticoliDescrEstesa ON tblArticoli.Cod_art=tblArticoliDescrEstesa.Cod_Art " & _
    "SET tblArticoli.Des_art_En = tblArticoliDescrEstesa.Des_Art_Estesa " & _
    "WHERE tblArticoli.Cod_art=tblArticoliDescrEstesa.Cod_Art; "

DoCmd.RunSQL strSQL
DoCmd.SetWarnings True
'dbs.Close
'Set dbs = Nothing
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdTestOdbc_Click()
Screen.MousePointer = vbHourGlass
Call OracleConnect
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'   Loop thru language table to set labels
    Dim dbs As Database, MySet As Recordset, strDefaultLanguage, strForm As String, hWndParent As Long, frmActive As Form

'   Look up the default language
    strDefaultLanguage = DLookup("[Selected_Language]", "tblDatiGenerali")
    strForm = Me.name
    'strForm = Screen.ActiveForm.Name

'   Assign the current database to the database variable
    Set dbs = CurrentDb
    
    Set MySet = dbs.OpenRecordset("select * from tblControlNames where strForm = '" & strForm & "'")
    
'   Testa se la select restituisce 0 records
    If MySet.RecordCount > 0 Then
        MySet.MoveFirst
        Do While Not MySet.EOF
            If MySet.Fields("strLanguage").Value = strDefaultLanguage Then
                Me(MySet.Fields("strControlName")).Caption = MySet.Fields("strControlCaption").Value
                
            
            End If
            MySet.MoveNext
        Loop
        MySet.Close
    End If
End Sub
