Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17456
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =1020
    Top =750
    Right =18480
    Bottom =9450
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x05c703a56ef2e340
    End
    Caption ="Portafoglio Ordini"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
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
        Begin Subform
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
        End
        Begin Tab
            Width =5103
            Height =3402
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
        Begin Page
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
        Begin Section
            CanGrow = NotDefault
            Height =8716
            BackColor =16777088
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =45
                    Top =600
                    Width =17385
                    Height =4995
                    Name ="fsubOcsamstSearch"
                    SourceObject ="Form.fsubOcsamstSearch"

                    LayoutCachedLeft =45
                    LayoutCachedTop =600
                    LayoutCachedWidth =17430
                    LayoutCachedHeight =5595
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Width =3855
                    Height =360
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    Name ="lblFormCaption"
                    Caption ="Monitor Portafoglio Ordini"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16755
                    Width =576
                    Height =576
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Command1"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"

                    LayoutCachedLeft =16755
                    LayoutCachedWidth =17331
                    LayoutCachedHeight =576
                End
                Begin Label
                    OverlapFlags =85
                    Left =480
                    Top =8160
                    Width =1830
                    Height =240
                    Name ="Label4"
                    Caption ="Scm Group - Spare Parts"
                    LayoutCachedLeft =480
                    LayoutCachedTop =8160
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =8400
                End
                Begin Subform
                    OverlapFlags =85
                    Left =45
                    Top =5640
                    Width =16935
                    Height =2250
                    TabIndex =2
                    Name ="fsubOcsaMstPO"
                    SourceObject ="Form.fsubOcsaMstPOSearch"

                    LayoutCachedLeft =45
                    LayoutCachedTop =5640
                    LayoutCachedWidth =16980
                    LayoutCachedHeight =7890
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16188
                    Width =567
                    Height =507
                    TabIndex =3
                    Name ="cmdCercaArticolo"
                    Caption ="Scorte"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Gestione Scorte"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000070809060202830ff4048504000000000 ,
                        0x0000000000000000000000000000000070584030705840ff7058403000000000 ,
                        0x00000000000000000000000000000000708090ff30b8f0ff101820ff40485040 ,
                        0x00000000000000000000000070584030705840fff0e8e0ffb0a090ff00000000 ,
                        0x0000000000000000000000000000000070809050708090ff30b8f0ff202840ff ,
                        0x404850400000000070584030705840fff0f0f0ffb0a090ff0000000000000000 ,
                        0x000000000000000000000000000000000000000070809050708090ff30b8f0ff ,
                        0x303850ff60505060705840fffff8f0ffb0a090ff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000070809050708090ff ,
                        0x40a8d0ff705840ffffffffffb0a090ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000070707070 ,
                        0x705840ffffffffffb0a090ff4048504000000000000000009070601090786040 ,
                        0x805850ff907860400000000000000000000000000000000070584030705840ff ,
                        0xffffffffb0a090ff30b8f0ff606070ff908880a00000000080605050806850ff ,
                        0xf0f0f0ff908070ff00000000000000000000000070584030705840ffffffffff ,
                        0xb0a090ff70809050708090ff70a0a0ff908070ff907060ff806050ff907860f0 ,
                        0xb09080ffb0a09090a0888070a08070ff806850ff907060ffffffffffb0a090ff ,
                        0x0000000000000000a09890a0a09080fff0f0f0ffe0e0d0ffd0c8c0ff907860e0 ,
                        0xb0988070b0a09020b0a090ffc0b0a0ffc0b0a0ffc0b0a0ff908070ff00000000 ,
                        0x0000000000000000b0a09040c0a090ffffffffffffffffe0f0e0e0ffb09080c0 ,
                        0x0000000000000000b0a090ffffffff20c0b0a030c0b0a0ffa08070ff00000000 ,
                        0x000000000000000090807050a08870fffff8fff0f0e0e0f0c0a090f0b0a09030 ,
                        0x00000000000000000000000000000000fff8ff40d0b8b0ffc0a8a0ff00000000 ,
                        0x90786020907860e0907060ffb0a8a0f0c0a8a0e0c0a090b0b0a0903000000000 ,
                        0x00000000000000000000000000000000b0a090ffb0a090ffb0a0905000000000 ,
                        0xc0a8a0ffc0a090ffd0b0a0ffc0b0a0ffb0a09050000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =16188
                    LayoutCachedWidth =16755
                    LayoutCachedHeight =507
                    PictureCaptionArrangement =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15795
                    Width =381
                    Height =366
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    ForeColor =10040115
                    Name ="cmdHelp"
                    Caption ="Help"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Composizione automatica"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000e0e8e000e0c8b000 ,
                        0xe0d8d000e0d0c010e0d0c010d0d0c010d0d0c000d0d0d000e0e0e00000000000 ,
                        0x0000000000000000000000000000000000000000f0e8e0009068303080582080 ,
                        0x905010c0804820e0804820c0804810b06040108050381030d0c8c01000000000 ,
                        0x000000000000000000000000e0780000e0a05010a0683070c08860f0e0c8b0ff ,
                        0xf0f0f0fffffffffffffffffff0f0f0ffe0c8c0ffa07850c040301060d0c8c010 ,
                        0xe0d8d0000000000000000000e0882000b0703070e0a880fffff0e0ffe0b8a0ff ,
                        0xd08050ffc05820ffc05820ffd08050ffe0b8a0fff0e8e0ffb09070f050301060 ,
                        0xd0c8c000e0e0e00000000000b0783030d09870f0fff0e0ffe0a890ffc05010ff ,
                        0xc05010ffe0a890ffffffffffb04810ffb04810ffd0a080fff0f0e0ffa07050d0 ,
                        0x50381030d0d0d000f0f0f000b0784080f0d8c0fff0c8b0ffe05820ffd05810ff ,
                        0xd05010ffe08050ffe0a880ffc05010ffb04810ffb04810ffe0b8a0ffe0c8c0ff ,
                        0x50401080d0d0d010f0f0f000d08040e0fff8f0fff09870fff06020ffe05820ff ,
                        0xe05820fff0a890ffffffffffd05010ffc05010ffb05010ffc07850fff0f0f0ff ,
                        0x804020c0e0d0c000f0f0f000d08040f0ffffffffff7840ffff6830fff06820ff ,
                        0xf06020fff08850fffffffffff0c0b0ffc05820ffb05010ffb05820ffffffffff ,
                        0x804820e0e0d0c010f0f0f000d08850f0ffffffffff8050ffff7030ffff6830ff ,
                        0xff6830ffff6820fff09060fffff8f0fff0d8c0ffc05020ffc05820ffffffffff ,
                        0x804820e0e0d8d010f0f0f000d08050c0fff8f0ffffa880ffff7040ffff8850ff ,
                        0xffb090ffff7030fff06820fff09070fffffffffff08050ffd08860fffff0f0ff ,
                        0x805820b0e0d8d010f0f0f000c0804070f0d8c0ffffd0c0ffff7840ffff9870ff ,
                        0xffffffffffc8b0ffff9060ffffc8b0fffff8f0fff07840fff0c8b0ffe0c8b0ff ,
                        0x90602070e0c8b00000000000c0884030e0a070f0fff8f0ffffc0a0ffff7840ff ,
                        0xffb8a0fffff8f0fffffffffffff0e0ffff9870fff0b8a0fffff0e0ffc08850e0 ,
                        0xa0682030f0e8e0000000000000000000c0884060e0b8a0f0fff8f0ffffd0c0ff ,
                        0xffa880ffff8850ffff8850ffffa880fff0d0c0fffff0e0ffd0a880f0a0683060 ,
                        0xe0c0a00000000000000000000000000000000000c0884060e0a070f0f0d8c0ff ,
                        0xfff8f0fffffffffffffffffffff8f0fff0d8c0ffc09060e0a0703050f0b89000 ,
                        0x0000000000000000000000000000000000000000f0f0f000c0884030c0804070 ,
                        0xe0a070c0d09870e0d09860f0d09870d0b0784070b0784020f0e8f00000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0f0f000f0f0f000f0f0f000f0f0f000f0f0f00000000000f0f0f00000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =15795
                    LayoutCachedWidth =16176
                    LayoutCachedHeight =366
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =8580
                    Top =7980
                    Width =615
                    Height =615
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =13209
                    Name ="cmdExport2txt"
                    Caption ="&Export Dati"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Esporta su fileTCV"
                    Picture ="ExportSavedExports.BMP"
                    UnicodeAccessKey =69
                    ImageData = Begin
                        0x424d360c00000000000036000000280000002000000020000000010018000000 ,
                        0x0000000c000000000000000000000000000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000000000ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00ffffffffffffffffffffffffffffffffffffffffff578bba5e88ad5e88ad5f ,
                        0x84ad5f85ad5c82a95981a65a8ab45890b55a91b65f8eb86593b76090b85b8cb2 ,
                        0x588aae5888ae5889b1588ab45a8bb55895be5490ba3958780000000000000000 ,
                        0x00ffffffffffffffffffffffffffffffffffffffffff6e97c1709fc69ce1fc61 ,
                        0xc5ef60c4f060c4f05fc4f060c4f061c4ef67c7f079cdf28ad4f482d1f370caf1 ,
                        0x64c6f060c4f060c4f05fc4f060c4ef5fc4f060c4f053a5cf4863800000000000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffff799bc579b2d4a3e3fc7c ,
                        0xd1f662c6f162c5f062c6f063c5f166c7f174ccf399daf6944c2ab1d4e393d7f5 ,
                        0x74ccf366c6f162c6f162c5f063c5f162c5f162c5f160c3ee4975a2151a1f0000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffff799dc784c7e37eb4d6aa ,
                        0xe6fe68c9f365c7f167c8f26ccaf373ccf389d5f5b8e5f9944c2a944c2ac0e0ef ,
                        0x97d9f677cdf369c9f265c7f265c7f265c7f165c8f265c7f259afd94b6d970000 ,
                        0x00000000000000ffffffffffffffffffffffffffffff7a9dc89ddbf273a7caa4 ,
                        0xe4ff8ddafa6ccaf376cef585d3f59ad9f5b3ddefd5e6ed984f2db66541954d2b ,
                        0xc3dfeb99daf779cff46ccaf368c9f369c9f368c9f469c9f368c8f24b7cab344b ,
                        0x5b000000000000000000ffffffffffffffffffffffff7b9fc9a6eafe74bddd7b ,
                        0xb3d5ace7ff7ed3f890d8f7b0cad5bcada6b18670954d2bb36540c16f49b96944 ,
                        0x964d2bc1dbe594d9f877cff56dcbf46bcbf46bcbf56ccbf56bcbf461bde54c72 ,
                        0x9f000707000000000000ffffffffffffffffffffffff7ca1cca2ebff77cdeb75 ,
                        0xa9c7a4e4feaee6fdb4d4e0b48068a45733bd6b46d27950d77c54d77c54d27e58 ,
                        0xc46f4b964d2bafd8ea83d4f772cef66fcdf66fcdf66fcdf66fcdf66fcdf6538d ,
                        0xb645657b0000000000000000000000000000000000007ca3cdaeeeff8cdff975 ,
                        0xbedd85b5d3c7eefec5b3aab25b37e18e67f0a47ff0a27cf79e75f7996cf0966c ,
                        0xda8d68c2724bb0d2df86d6f875d0f772cff772cef772cff772cff772cff76cc8 ,
                        0xf04e79a60b11170000000000000000000000000000007ea5cfb6eeff9ae9ff7b ,
                        0xcdeb90b9d2cee4ecc47e5ff1a482fec4a9ffc4a6ffc4a6fdc2a3fdb896fcc1a2 ,
                        0xc2724bccd6d99fdefa81d5f976d1f875d0f975d1f875d1f975d1f875d1f975d1 ,
                        0xf860a3cd6a7785806d59806d59806d590000000000007fa7d2b8f1ff9ae9ff8a ,
                        0xdef89fd4eaccb3a8dc8a69f9bb9fcd9074d8bcaee8d7cfbb6a44fcc1a2c2724b ,
                        0xcfd6d9a8e2fb89d8fb7ad3f977d2f977d2fa78d2f977d2f977d2fa77d1fa77d2 ,
                        0xf976d1f8527eb0e3e0def8f0ea806d5900000000000080a9d3b9f0ff9deaff98 ,
                        0xe9feafe1f2d2977cf9bb9fc78d74d4dde1c5ecfdd2f0fdc2724bc2724bd0d7d9 ,
                        0xa9e3fc8ad9fb7dd4fa7ad3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3 ,
                        0xfa79d3fa6fbde47995b5f8f0eb806d5900000000000081abd5b8f0ffa1ebff9e ,
                        0xeafeb9eaf9d0896acf8661dbe5e7bbeafe9cdffbafe5fcc2724bc7cfd3a8e2fc ,
                        0x8bd9fb7dd4fa7ad3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3 ,
                        0xfa79d3fa79d3fa5186bbc8ced4806d5900000000000082acd7b9f0ffa4ecffa3 ,
                        0xebffbdf0fecd8566c9b7acbddbf4b2e6fdb5eaffbbecffc4eeffc1edffb7ebff ,
                        0xb1e9ffafe8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8 ,
                        0xffaee8ffaee8ff7fc2e46295c0806d5900000000000083afd9bcf1ffa7edffae ,
                        0xeeffb6effedb9e81ccf0f8a7e0fa9ebbd89ab8d79bbad89dbbd99bbad897b8d7 ,
                        0x93b6d692b5d791b6d790b4d78fb4d78eb5d88eb4d77eaedf78aee578aee578ae ,
                        0xe570a7de68a1d65e95cf629bd2806d5900000000000084b1dcbdf2ffa9edffac ,
                        0xeeffb4f1ffbcf1ffb2f0ffa4ecffc8b5a5fbf9f6fbf8f5fbf8f5fbf6f4faf6f4 ,
                        0xfbf5f2c7b4a4fbf9f6fbf8f5fbf8f5fbf6f4faf6f4fbf5f2c1ad9efaf4effaf3 ,
                        0xeefaf2eef9f2edf8f1edf9f1ec806d5900000000000085b2ddbdf2ffaaefffa9 ,
                        0xeeffabeeffadeeffa9edffa2ebffc6b4a5fcf9f7fbf8f6fcf8f5fcf8f5fbf7f4 ,
                        0xfbf6f3c6b4a5fcf9f7fbf8f6fcf8f5fcf8f5fbf7f4fbf6f3c1ae9ffaf4f0f9f3 ,
                        0xeffaf2eef9f2eef9f1edf9f1ed806d5900000000000086b3debef2ffa8efffa7 ,
                        0xefffa9efffadeeffaaeeffa5ebffc7b4a5fcf9f7fbf8f7fbf8f6fcf8f5fbf7f4 ,
                        0xfbf7f4c7b4a5fcf9f7fbf8f7fbf8f6fcf8f5fbf7f4fbf7f4c1ae9efaf4f0faf3 ,
                        0xf0faf3effaf3eef9f2edf9f1ec806d5900000000000086b4dfb3e6fab1efffa8 ,
                        0xeeffacefffb0f0ffaceeffa7efffc8b4a5fcfaf7fcf9f6fcf8f6fbf8f6fbf7f5 ,
                        0xfbf7f5c8b4a5fcfaf7fcf9f6fcf8f6fbf8f6fbf7f5fbf7f5c2af9ffaf4f1faf4 ,
                        0xf1f9f3effaf2eff9f2eef9f2ed806d5900000000000078acdb8dbee5b8ebfcbf ,
                        0xf2ffc1f2ffc1f3ffc2f2ffbff1ffc7b5a6c7b4a6c6b4a5c6b3a4c5b2a3c5b1a2 ,
                        0xc5b1a2c7b5a6c7b4a6c6b4a5c6b3a4c5b2a3c5b1a2c5b1a2c3ae9fc1ae9fc1ad ,
                        0x9fc0ac9dc0ac9ec0ac9cbfab9c806d59000000000000ffffff93bee88cb8e38c ,
                        0xb8e38bb8e38bb8e389b8e388b7e3c8b5a6fefefefefdfdfdfdfdfdfcfcfdfdfc ,
                        0xfdfbfbc8b5a6fdf9f9fcfaf8fbf9f7fbf9f6fbf8f6fbf7f5c3afa0fbf4f1faf4 ,
                        0xf0faf4f0f9f3eff9f3eff9f2ee806d59000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffc8b6a6fefdfefefdfdfefdfdfefdfdfefdfc ,
                        0xfdfcfcc8b6a6fcfaf9fcf9f8fcf9f8fcf8f7fcf8f6fbf8f6c3afa0faf5f2faf5 ,
                        0xf2faf4f1f9f4f0f9f3eff9f2ef806d59000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffc8b6a7fefdfefefefdfefefdfefdfdfefcfc ,
                        0xfdfcfcc8b6a7fcfbf9fcfaf8fcf9f8fcf9f8fcf9f7fcf8f6c3b0a1faf5f2faf5 ,
                        0xf1faf4f1faf4f1f9f3f0f9f3ef8b7966000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffc9b6a7fefefefefdfefefefefefefdfdfdfd ,
                        0xfdfdfdc9b6a7fdfbf9fdfaf9fcfaf8fcfaf8fcf8f7fcf9f7c3b0a1fbf6f3faf5 ,
                        0xf2faf5f2faf4f1faf4f0f9f4ef988575000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb88971b27c61b17b60b17a5fb17a5eb0795d ,
                        0xb0785bad7356ad7255ac7153ac7152ab6f51aa6e50aa6d4ea86949a76848a767 ,
                        0x46a66646a66544a56543a56343a4603f000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb88971d7baabd7baabd7baabd7baabd7baab ,
                        0xd7baabae7456d7b9aad6b9aad6b8a9d6b8a9d5b7a7d5b6a7a86949d1af9ed0ad ,
                        0x9cceaa9acca897caa694c9a391a4613f000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb88971d7baabd7baabd7baabd7baabd7baab ,
                        0xd7baabad7457d7b9aad7b9aad6b9aad6b8a9d6b8a9d5b7a7a8694ad2b1a0d1af ,
                        0x9ed0ad9bceaa98cca896caa693a4613f000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb88971d7baabd7baabd7baabd7baabd7baab ,
                        0xd7baabae7457d7baabd7b9aad6b9aad6b9a9d6b8a9d6b8a8a8694ad2b2a2d1b1 ,
                        0xa0d0af9dcfad9bceaa98cca895a46140000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb88971b88870b88870b78870b7876eb6866e ,
                        0xb6866db58469b58269b58167b48267b38066b37f64b27f63b17b5fb07b5eb07a ,
                        0x5db07a5daf795caf785baf785aae7658000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff
                    End

                    LayoutCachedLeft =8580
                    LayoutCachedTop =7980
                    LayoutCachedWidth =9195
                    LayoutCachedHeight =8595
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9880
                    Top =8029
                    Width =525
                    Height =525
                    TabIndex =6
                    Name ="Comando35"
                    Caption ="Apri Dir"
                    ControlTipText ="Apri direttorio export dati"
                    Picture ="FileOpenDatabase.BMP"
                    ImageData = Begin
                        0x424d360c00000000000036000000280000002000000020000000010018000000 ,
                        0x0000000c000000000000000000000000000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff00000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00ffffffffffffffffffffffffffffffffffffffffffffffff00000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffffffffff588cbd5e88ad5e ,
                        0x88ad5f84ad5f85ad5c82a95981a65a8ab45890b55990b55889b55b8cb2588ab4 ,
                        0x588ab1588aae5888ae5889b1588ab45a8bb55895be5490ba426489141a210000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffffffffff6e97c1709fc69c ,
                        0xe1fc61c5ef60c4f060c4f05fc4f060c4f060c4ef60c4ef60c4f060c4f05fc4f0 ,
                        0x60c4f060c5f060c4f060c4f05fc4f060c4ef5fc4f060c4f053a5cf5575970000 ,
                        0x00000000000000ffffffffffffffffffffffffffffffffffff799bc579b2d4a3 ,
                        0xe3fc7cd1f662c6f162c5f062c6f062c5f162c6f162c6f162c6f162c6f062c5f1 ,
                        0x62c5f162c5f162c5f162c6f162c5f063c5f162c5f162c5f160c3ee4975a24557 ,
                        0x66000000000000000000ffffffffffffffffffffffffffffff799dc784c7e37e ,
                        0xb4d6aae6fe68c9f365c7f165c7f265c7f265c7f265c8f265c7f265c7f265c7f2 ,
                        0x65c7f265c7f265c7f265c8f265c7f265c7f265c7f165c8f265c7f259afd95074 ,
                        0xa023353b0c0c0c000000ffffffffffffffffffffffffffffff7a9dc89ddbf273 ,
                        0xa7caa4e4ff8ddafa68c9f369c9f468c9f368c9f368c9f368c9f368c9f368c9f4 ,
                        0x68c9f368c9f368c9f369c9f368c9f368c9f369c9f368c9f469c9f368c8f24b7c ,
                        0xab5379901d2b33131313000000ffffffffffffffffffffffff7b9fc9a6eafe74 ,
                        0xbddd7bb3d5abe7ff73cff76bcbf56ccbf46bcbf46ccbf46ccbf46bcbf46bcbf5 ,
                        0x6bcbf46bcbf56bcbf56ccbf56bcbf46ccbf46bcbf46bcbf56ccbf56bcbf461bd ,
                        0xe54e75a2314e530b1616000000ffffffffffffffffffffffff7ca1cca2ebff77 ,
                        0xcdeb74a8c79fe3fe9be0fd6fcdf56fcdf66ecdf66fcdf66fcdf66fcdf66fcdf5 ,
                        0x6ecdf66ecdf66fcdf66fcdf66fcdf66fcdf66fcdf66fcdf66fcdf66fcdf66fcd ,
                        0xf6538db6577e99152424000000000000ffffffffffffffffff7ca3cdaeeeff8c ,
                        0xdff973bddc78adcea9e6ff7fd5fa72cff772cff772cff872cff872cff872cff7 ,
                        0x72cff772cff772cff872cff772cff772cff772cff772cef772cff772cff772cf ,
                        0xf76cc8f04e79a73e53600a0a0a000000000000ffffffffffff7ea5cfb6eeff9a ,
                        0xe9ff76cbea75a8c79be2ffa5e5ff76d1f975d1f875d1f875d1f875d0f975d1f9 ,
                        0x75d1f875d1f875d1f975d0f875d1f875d1f875d0f975d1f875d1f975d1f875d1 ,
                        0xf975d1f85fa3cd5a7ca21d1d24000000000000ffffffffffff7fa7d2b8f1ff99 ,
                        0xe9ff81dbf774c1e073b3d8a5e5ff8adafc78d2f977d2fa78d2f977d2fa77d2fa ,
                        0x77d2f977d2fa78d2f977d2fa77d2f977d2f977d2fa78d2f977d2f977d2fa77d1 ,
                        0xfa77d2f976d1f8507daf4c697b111a1a000000000000ffffff80a9d3b9f0ff9c ,
                        0xeaff8ce6fe78ccea72bee289cef4a9e6ff7bd4fb79d3fa79d3fa79d3fa79d3fa ,
                        0x79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3 ,
                        0xfa79d3fa79d3fa6ebce35a7ea8314248162121000000ffffff81abd5b8f0ff9f ,
                        0xebff90e7fe82d9f473caea6fb3dda1e4ff92ddfd79d3fa79d3fa79d3fa79d3fa ,
                        0x79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3fa79d3 ,
                        0xfa79d3fa79d3fa79d3fa5085ba577e9c272f2f000000ffffff82acd7b9f0ffa3 ,
                        0xecff95e8ff8be5fd7cd2ef70c7e890c3eca4e2fdaee8ffaee8ffaee8ffaee8ff ,
                        0xaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8ffaee8 ,
                        0xffaee8ffaee8ffaee8ff79bfe34785b8474747000000ffffff83afd9bcf1ffa6 ,
                        0xedffa5ecff93e8fe9ceaff9febff8fd7f88cb8e38cb8e38cb8e38ab7e389b7e3 ,
                        0x88b7e386b7e385b6e484b7e483b5e482b5e481b6e580b5e578aee578aee578ae ,
                        0xe578aee570a7de68a1d65793d25197db555555000000ffffff84b1dcbdf2ffa8 ,
                        0xedffa8edffa7eeffa6ecffa0ecff9ceaff94e7ff8ce4ff86e3ff81e3ff7de2ff ,
                        0x78e3ff75e0ff73dcfa71dcfa6cd9f86cd7f669d4f669d1f166cdf05995c40000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffffffffff85b2ddbdf2ffaa ,
                        0xefffa8eeffa8edffa8edffa5ecffa1ebff99e9ff93e8ff8de6ff8ce6ff97eaff ,
                        0x95e9ff90e7ff8de7fe89e4fd88e1fa85defb83dcf881d9f77cd4f6588ac40000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffffffffff86b3debef2ffa8 ,
                        0xefffa7efffa9efffaceeffa9eeffa5ebff9febff9be9ff9ce8ffa3e3fb8dc6f0 ,
                        0x81b6e581b6e580b5e57fb5e57fb5e57cb4e478aee477ade175b1e574b9df0000 ,
                        0x00000000000000000000000000ffffffffffffffffffffffff86b4dfb3e6fab1 ,
                        0xefffa8eeffacefffb0f0ffaceeffa7efffa1ecffa0ebffa5e2fa8dbde76f9cc4 ,
                        0x000000000000ffffffffffffffffffffffffffffffffffffffffff0000000000 ,
                        0x00000000000000000000000000ffffffffffffffffffffffff79addd8dbee5b8 ,
                        0xebfcbff2ffc1f2ffc1f3ffc2f2ffbff1ffb6f0fface6fa90c1e97aa6d2000000 ,
                        0x000000ffffffffffffffffff000000000000000000ffffffffffff8269578063 ,
                        0x4f7f5d497f5a44000000000000ffffffffffffffffffffffffffffff93bee88c ,
                        0xb8e38cb8e38bb8e38bb8e389b8e388b7e387b7e486b7e482aedb000000000000 ,
                        0xffffffffffffffffffffffff0000000000000000000000000000000000007c61 ,
                        0x508064537f5e4a000000000000ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff7f552a7e64526955460000000000000000006853458064 ,
                        0x53775d4e816552000000000000ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff7d63527f6453755b4c7058496f57487e62527a60 ,
                        0x4f000000857060000000000000ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff7d61508064538165547f6453785d4d0000 ,
                        0x00000000ffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff
                    End

                    LayoutCachedLeft =9880
                    LayoutCachedTop =8029
                    LayoutCachedWidth =10405
                    LayoutCachedHeight =8554
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

Private Sub cmdCercaArticolo_Click()
'wizard code to open form, linking the lstMember ID number
    'to that on the Membership form
'this button then deleted as just wanted the code procedure
'called from lstMember_DblClick
On Error GoTo Err_cmdOpenMember_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim rst As DAO.Recordset
    stLinkCriteria = "[COD_ART] = '" & Forms!frmOcsaMstSearch!fsubOcsamstSearch.Form.txtCod_Art & "'"

    stDocName = "frmParts1"
    
    ' Open a recordset to see if any rows returned with this filter
    Set rst = DBEngine(0)(0).OpenRecordset("SELECT * FROM tblarticoli WHERE " & stLinkCriteria)
    ' See if found none
    If rst.RecordCount = 0 Then
        MsgBox "Nessun codice trovato.", vbInformation, gstrAppTitle
        ' Clean up recordset
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpenMember_Click:
    Exit Sub

Err_cmdOpenMember_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenMember_Click
End Sub

Private Sub cmdClose_Click()
    ' User wants to save any pending work and exit
    ' Call the common save routine - returns False if save failed
    ' Close me
    DoCmd.Close acForm, Me.name
End Sub
Private Sub cmdExport2txt_Click()
' Esporta Portafoglio Ordini
Dim rst As Recordset
Dim Varwhere As Variant
    ' Dim rst As Object
    
   
     ' Open a recordset to see if any rows returned with this filter
     ' Cerca se trova niente
     
     'Set rst = Forms!frmOcsaMstSearch!fsubOcsamstSearch.Form.RecordsetClone
      
    Set rst = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryCOrdersSearch")
        If rst.RecordCount = 0 Then
            MsgBox "Non ho trovato nessun ordine.", vbInformation, gstrAppTitle
            ' Clean up recordset
            rst.Close
            Set rst = Nothing
            
            Exit Sub
        End If
    
'/// Inizia a creare File Csv

Dim z As Long
Dim q As Boolean


q = False

' Dim conCurrent As ADODB.Connection
Dim Db As DAO.Database
Dim rstOutput As New ADODB.Recordset
Dim objField As ADODB.Field
Dim intFile As Integer
Dim strSQL As String, strDataLine As String
Dim filenm As String
Dim i As Integer
' Costruzione File da aprire
intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "PortafoglioOrdini_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
    & "PortafoglioOrdini_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
            & "PortafoglioOrdini_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "PortafoglioOrdini_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrAppTitle) Then
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Pr" & Chr(9) & "N_Doc" & Chr(9) & "Data_Ordine" & Chr(9) & "N_Cliente" & Chr(9) & "Rag_Soc" & Chr(9) & "Codice" & Chr(9) & _
                "Descrizione" & Chr(9) & "Qta_Ordine" & Chr(9) & "Data_Consegna" & Chr(9) & "Qtà_Spedita" & Chr(9) & _
                "Giacenza" & Chr(9) & "Impegnato" & Chr(9) & "Giacenza_Stefani" & Chr(9) & "N_Doc_Acq" & Chr(9) & _
                "Cod_F" & Chr(9) & "Rag_Soc_Fornitore" & Chr(9) & "Data_Ord_F" & Chr(9) & _
                "Qta_Residuo" & Chr(9) & "Data_Consegna"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
    If Not rst.EOF And Not rst.BOF Then
    rst.MoveFirst
    End If
    While Not rst.EOF
    strDataLine = rst.Fields("Liv_Urgenza").Value & Chr(9) & _
                    rst.Fields("Numero_doc").Value & Chr(9) & _
                    rst.Fields("Data_Ordine").Value & Chr(9) & _
                    rst.Fields("Cod_Cli").Value & Chr(9) & _
                    rst.Fields("Ds_Rag_soc").Value & Chr(9) & _
                    rst.Fields("Cod_Art").Value & Chr(9) & _
                    rst.Fields("Descrizione").Value & Chr(9) & _
                    rst.Fields("Qta_Ord_umv").Value & Chr(9) & _
                    rst.Fields("Data_Prev_Cons").Value & Chr(9) & _
                    rst.Fields("Qta_Cons_umv").Value & Chr(9) & _
                    rst.Fields("DispSp").Value & Chr(9) & _
                    rst.Fields("Impegnato").Value & Chr(9) & _
                    rst.Fields("DispAh").Value & Chr(9) & _
                    rst.Fields("Ord_Acq").Value & Chr(9) & _
                    rst.Fields("Cod_Forn").Value & Chr(9) & _
                    rst.Fields("Rag_Soc_Forn").Value & Chr(9) & _
                    rst.Fields("Data_Ordine").Value & Chr(9) & _
                    rst.Fields("Qta_Residua").Value & Chr(9) & _
                    rst.Fields("Data_Ric").Value
    Print #intFile, strDataLine
    rst.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "PortafoglioOrdini" & i & ".xls")
    rst.Close
    Set rst = Nothing
    Set Db = Nothing
    Close #intFile

End If

End Sub
Private Sub cmdHelp_Click()
Dim FormHelpId As Long
Dim curForm As Form
    'Set the curForm variable to the currently active form.
    Set curForm = Screen.ActiveForm
    FormHelpId = 180
    'Call the function to start the Help file, passing it the name of the
    'Help file and context-id.
    Show_Help FormHelpFile, FormHelpId
End Sub




Private Sub txtNumero_Doc_AfterUpdate()
    [Forms]![frmOcsaMst]![fsubOcsamst].Requery
'    Me!txtRag_Soc.Enabled = False
'    Me!txtRag_Soc = Null
'If Not IsNull(txtNumero_Doc) Then
'    Me!txtRag_Soc.Enabled = True
'    Me!txtRag_Soc.Enabled = False
'End If
End Sub
