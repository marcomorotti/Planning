Version =20
VersionRequired =20
Begin Form
    CloseButton = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =0
    GridY =10
    Width =5797
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =2640
    Top =1128
    Right =8976
    Bottom =7488
    RecSrcDt = Begin
        0xd4c266956ee4e340
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
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
            SizeMode =3
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
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
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
            FELineBreak = NotDefault
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
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
        Begin ListBox
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
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
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
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
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =591
            BackColor =12311007
            Name ="IntestazioneMaschera"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =15
                    Top =60
                    Width =3585
                    Height =495
                    FontSize =18
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Auto_Title0"
                    Caption ="Articoli dati aggiuntivi"
                    FontName ="Segoe UI"
                    GridlineColor =-2147483616
                    HorizontalAnchor =2
                    LayoutCachedLeft =15
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =555
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =4650
                    Width =1077
                    Height =450
                    FontSize =12
                    FontWeight =700
                    ForeColor =255
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Save e Close the window."
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4650
                    LayoutCachedWidth =5727
                    LayoutCachedHeight =450
                End
            End
        End
        Begin Section
            Height =5782
            BackColor =12311007
            Name ="Corpo"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =345
                    Width =3630
                    Height =359
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =255
                    Name ="txtID_ArticoliStato"
                    RowGroup =1
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =345
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =704
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =120
                            Top =345
                            Width =1904
                            Height =359
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =255
                            Name ="Etichetta1"
                            Caption ="ID_ArticoliStato:"
                            RowGroup =1
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =345
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =704
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =779
                    Width =3630
                    Height =359
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtCod_Art"
                    RowGroup =2
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =779
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =1138
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =779
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta3"
                            Caption ="Codice Articolo:"
                            RowGroup =2
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =779
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =1138
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =2905
                    Width =3630
                    Height =1276
                    TabIndex =6
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtNote"
                    RowGroup =4
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =2905
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =4181
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =2905
                            Width =1904
                            Height =1276
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta7"
                            Caption ="Note:"
                            RowGroup =4
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =2905
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =4181
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =2037
                    Width =3630
                    Height =359
                    ColumnWidth =2265
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtCod_Art_Correlato"
                    RowGroup =5
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =2037
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =2396
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =2037
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta9"
                            Caption ="Articolo Correlato:"
                            RowGroup =5
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =2037
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =2396
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2084
                    Top =1647
                    Width =3630
                    Height =315
                    ColumnWidth =1650
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="txtID_StatoArticolo"
                    RowSourceType ="Table/Query"
                    RowSource ="select NullId as ID_StatoArticolo, NullId as Stato, NullId as SequenzaStato from"
                        " tblDummy UNION SELECT tblStatoArticolo.ID_StatoArticolo, tblStatoArticolo.Stato"
                        ", tblStatoArticolo.SequenzaStato FROM tblStatoArticolo order by SequenzaStato;"
                    ColumnWidths ="0;1134"
                    RowGroup =3
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =2084
                    LayoutCachedTop =1647
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =1962
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =1647
                            Width =1904
                            Height =315
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta5"
                            Caption ="Stato Articolo:"
                            RowGroup =3
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =1647
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =1962
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =2471
                    Width =3630
                    Height =359
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtScortaSicurezzaForzata"
                    RowGroup =6
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =2471
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =2830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =1
                            Left =120
                            Top =2471
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta11"
                            Caption ="Scorta S Forzata:"
                            RowGroup =6
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =2471
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =2830
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =4260
                    Width =3630
                    Height =359
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtLotto_Min"
                    RowGroup =7
                    GroupTable =8

                    LayoutCachedLeft =2084
                    LayoutCachedTop =4260
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =4619
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =4260
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta13"
                            Caption ="Lotto Acq Min:"
                            RowGroup =7
                            GroupTable =8
                            LayoutCachedLeft =120
                            LayoutCachedTop =4260
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =4619
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =4740
                    Width =3630
                    Height =359
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtLotto_Multiplo"
                    RowGroup =8
                    GroupTable =9

                    LayoutCachedLeft =2084
                    LayoutCachedTop =4740
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =5099
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =4740
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta17"
                            Caption ="Lotto Acq Multiplo:"
                            RowGroup =8
                            GroupTable =9
                            LayoutCachedLeft =120
                            LayoutCachedTop =4740
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =5099
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2084
                    Top =1213
                    Width =3630
                    Height =359
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtDes_Art"
                    RowGroup =9
                    GroupTable =6

                    LayoutCachedLeft =2084
                    LayoutCachedTop =1213
                    LayoutCachedWidth =5714
                    LayoutCachedHeight =1572
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =1
                            Left =120
                            Top =1213
                            Width =1904
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Etichetta14"
                            Caption ="Descrizione:"
                            RowGroup =9
                            GroupTable =6
                            LayoutCachedLeft =120
                            LayoutCachedTop =1213
                            LayoutCachedWidth =2024
                            LayoutCachedHeight =1572
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PièDiPaginaMaschera"
            AutoHeight =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
     DoCmd.Close acForm, Me.name
End Sub

Public Sub cmdSave_Click()
    Dim strSQL As String
    Dim ScortaSicurezzaForzata As Variant
    Dim Record As Integer
    If IsNull(Me!txtID_StatoArticolo) Then
        Record = 0
    Else: Record = 1
    End If
    If IsNull(Me!txtCod_Art_Correlato) Then
        Record = Record
    Else: Record = Record + 1
    End If
    If IsNull(Me!txtScortaSicurezzaForzata) Then
        Record = Record
    Else: Record = Record + 1
    End If
    If IsNull(Me!txtNote) Then
        Record = Record
    Else: Record = Record + 1
    End If
    If IsNull(Me!txtLotto_Min) Then
        Record = Record
    Else: Record = Record + 1
    End If
    If IsNull(Me!txtLotto_Multiplo) Then
        Record = Record
    Else: Record = Record + 1
    End If
    If Record > 0 Then
        ScortaSicurezzaForzata = IIf(IsNull(Me!txtScortaSicurezzaForzata), "Null", Me!txtScortaSicurezzaForzata)
        Lotto_min = IIf(IsNull(Me!txtLotto_Min), "Null", Me!txtLotto_Min)
        Lotto_multiplo = IIf(IsNull(Me!txtLotto_Multiplo), "Null", Me!txtLotto_Multiplo)
        ID_StatoArticolo = IIf(IsNull(Me!txtID_StatoArticolo), "Null", Me!txtID_StatoArticolo)
        strSQL = "INSERT INTO tblArticoliStato ([Cod_Art]," & _
                 "[ID_StatoArticolo], [Note], [Cod_Art_Correlato], [Data_Modifica], [ScortaSicurezzaForzata], [Lotto_min], [Lotto_multiplo], [DES_ART]) " & _
                 "VALUES ('" & Me!txtCod_Art & "'," & _
                 ID_StatoArticolo & ", " & _
                 "'" & fncSQLStr(Me!txtNote) & "'," & _
                 "'" & Me!txtCod_Art_Correlato & "'," & _
                 "#" & Format(Date, "mm/dd/yyyy") & "#, " & _
                 ScortaSicurezzaForzata & "," & _
                 Lotto_min & "," & _
                 Lotto_multiplo & "," & _
                 "'" & fncSQLStr(Me!txtDes_art) & "'" & ")"
        CurrentDb.Execute strSQL, dbFailOnError
    End If

    If IsLoaded("frmParts1") Then
        [Forms]![frmParts1].Requery
    End If

    DoCmd.Close acForm, Me.name
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.txtCod_Art = [Forms]![frmParts1]![txtCod_Art]
    Me.txtDes_art = [Forms]![frmParts1]![txtDes_art]
End Sub
