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
    Left =9810
    Top =4425
    Right =15075
    Bottom =5685
    HelpContextId =52
    RecSrcDt = Begin
        0xa3aa30f84e40e140
    End
    Caption ="Articoli Obsoleti"
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
                    OverlapFlags =85
                    TextAlign =1
                    Top =45
                    Width =2655
                    Height =405
                    FontSize =14
                    ForeColor =255
                    Name ="Text12"
                    Caption ="Anagrafica ClientiI"
                    LayoutCachedTop =45
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =450
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4654
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
                    ControlTipText ="Close the form"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4654
                    LayoutCachedWidth =5230
                    LayoutCachedHeight =576
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    TextFontFamily =34
                    Left =4077
                    Width =577
                    Height =577
                    TabIndex =1
                    Name ="cmdClienti"
                    Caption ="E&xcel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Export ..."
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

                    LayoutCachedLeft =4077
                    LayoutCachedWidth =4654
                    LayoutCachedHeight =577
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3501
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
                    ControlTipText ="Apre direttorio export"
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

                    LayoutCachedLeft =3501
                    LayoutCachedWidth =4077
                    LayoutCachedHeight =576
                End
            End
        End
        Begin Section
            Height =708
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Width =1755
                    Height =600
                    ForeColor =255
                    Name ="Etichetta22"
                    Caption ="Premi tasto Excel  per esportare i dati"
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =600
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



Private Sub cmdClienti_Click()
    Call GenerateCsvClienti
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.name
End Sub

'Private Sub cmdHelpReportCost_Click()
'DoCmd.OpenForm "z Help Text for User", , , "zhID = 90"
'End Sub


'Private Sub cmdHelp_Click()
'Dim hwndHelp As Long
'Call HelpEntry
'End Sub
