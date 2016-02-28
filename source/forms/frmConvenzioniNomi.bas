Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =10155
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =2025
    Top =2430
    Right =12180
    Bottom =8565
    RecSrcDt = Begin
        0x4d4b221f0148e340
    End
    Caption ="Convenzione dei nomi tabelle"
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
        Begin Section
            Height =6150
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin Label
                    SpecialEffect =3
                    OverlapFlags =93
                    Width =10155
                    Height =5205
                    FontSize =12
                    ForeColor =16711680
                    Name ="lblSampleTables"
                    Caption ="Le tabelle di questo database e i campi, sono conformi ad uno standard\015\012\015"
                        "\012Nome              Prefisso        Esempio\015\012--------------     --------"
                        "------   ------------------ \015\012Tabelle            tbl                tblNom"
                        "eTabella\015\012Query             qry                qryNomeQuery\015\012QueryRi"
                        "cerca   qlk                qlkNomeQuery\015\012Query Dialogo  fdlg              "
                        "FileDialog\015\012Maschera         Frm              frmNomeMaschera\015\012Date "
                        "              99/99/0000;0\015\012\015\012Non usare speciali caratteri (#, @ , *"
                        ",  / , %) nei nomi degli oggetti. While you can get away with using them, doing "
                        "so is like playing with matches. Sooner or later someone is going to get burned."
                        "\015\012\015\012Non lasciare spazi tra le parole."
                    FontName ="Tahoma"
                    LayoutCachedWidth =10155
                    LayoutCachedHeight =7605
                End
                Begin Label
                    FontUnderline = NotDefault
                    SpecialEffect =1
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Top =4919
                    Width =10140
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BackColor =16711680
                    ForeColor =10092543
                    Name ="lblEmailGPCData"
                    Caption ="Email to Marco Morotti"
                    FontName ="Tahoma"
                    ControlTipText ="Email George at Grover Park Consulting"
                    HyperlinkAddress ="mailto:m.morotti@renco.it?subject=Here%20are%20my%20thoughts%20on%20your%20taxon"
                        "omy%20of%20Access%20Tables"
                    LayoutCachedTop =8234
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =8534
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =8040
                    Top =5535
                    Width =1290
                    Height =390
                    Name ="cmdDone"
                    Caption ="&Esci"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =69
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =8040
                    LayoutCachedTop =9210
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =9600
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
Option Compare Database
Option Explicit

Private Sub cmdDone_Click()
    ' Exit the application -
        DoCmd.Close acForm, Me.name
End Sub
