Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    PictureTiling = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =5
    PictureSizeMode =3
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =3195
    Top =1050
    Right =11775
    Bottom =6480
    DatasheetGridlinesColor =12632256
    PaintPalette = Begin
        0x000301000000000000000000
    End
    RecSrcDt = Begin
        0x00f075e70fe5e340
    End
    RecordSource ="SELECT DISTINCTROW tblTopics.strSampleCategory, tblTopics.strTopic, tblTopics.pk"
        "eyQNumber, tblTopics.strArticleTitle, tblTopics.memDescription, tblTopics.hrefUR"
        "L, tblTopics.strObjectName, tblTopics.intObjectType, tblTopics.strObjectsUsed FR"
        "OM tblTopics ORDER BY tblTopics.strSampleCategory, tblTopics.strTopic; "
    Caption ="Menu Report"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    PictureSizeMode =3
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BorderWidth =1
            FontWeight =700
            BackColor =12632256
            ForeColor =8404992
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
        Begin Line
            SpecialEffect =3
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
            SpecialEffect =3
            BackStyle =0
            OldBorderStyle =0
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
            FontSize =8
            ForeColor =128
            FontName ="Arial"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin CheckBox
            SpecialEffect =2
            BorderWidth =3
            LabelX =230
            LabelY =-30
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
        Begin OptionGroup
            SpecialEffect =3
            BackStyle =1
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
        Begin BoundObjectFrame
            SpecialEffect =2
            BorderColor =12632256
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
            BorderColor =12632256
            FontName ="Arial"
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
            BorderColor =12632256
            FontName ="Arial"
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
            BorderColor =12632256
            FontName ="Arial"
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
            SpecialEffect =3
            BorderColor =12632256
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
        Begin UnboundObjectFrame
            SpecialEffect =3
            BackStyle =0
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
        Begin ToggleButton
            FontSize =8
            ForeColor =128
            FontName ="Arial"
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
        Begin FormHeader
            Height =930
            BackColor =12311007
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =1080
                    Top =60
                    Width =3420
                    Height =420
                    FontSize =14
                    ForeColor =13209
                    Name ="lblAccess"
                    Caption ="Scm Gestione Scorte"
                    FontName ="Verdana"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =480
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =1590
                    Top =480
                    Width =4320
                    Height =420
                    FontSize =14
                    ForeColor =13209
                    Name ="Label26"
                    Caption ="Report"
                    FontName ="Verdana"
                End
                Begin UnboundObjectFrame
                    OverlapFlags =85
                    Width =955
                    Height =910
                    Name ="NonAssociatoOLE28"
                    OleData = Begin
                        0x00700000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000010000000000000000100000 ,
                        0x0200000001000000feffffff0000000000000000ffffffffffffffffffffffff ,
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
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdffffff06000000feffffff04000000050000003400000035000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000feffffff1300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x28000000290000002a0000002b0000002c0000002d0000002e0000002f000000 ,
                        0x30000000310000003200000033000000feffffff36000000fefffffffeffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff04000000141a020000000000c0000000 ,
                        0x00000046000000000000000000000000405a362fc346cc0103000000c0090000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000014000000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010100000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000076000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000018000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000030000002c050000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0xfefffffffefffffffeffffff1b0000001c0000001d0000001e000000feffffff ,
                        0x20000000210000002200000023000000240000002500000026000000feffffff ,
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
                        0xffffffff0100000200000000000000000000000000000000ae000000af000000 ,
                        0xb0000000b1000000b2000000b3000000b4000000b5000000b6000000b7000000 ,
                        0xb80000000100feff030a0000ffffffff141a020000000000c000000000000046 ,
                        0x1b00000044697365676e6f206469204d6963726f736f667420566973696f0012 ,
                        0x000000566973696f2031312e30205368617065730011000000566973696f2e44 ,
                        0x726177696e672e313100f439b2710000000000000000000000000000d7000000 ,
                        0xd8000000ffffffff030000000400000001000000ffffffff0200000000000000 ,
                        0xa30500003d080000be0400000100090000035f02000006005700000000000400 ,
                        0x000003010800050000000b0200000000050000000c0289036802030000001e00 ,
                        0x07000000fc020000ffffff000000040000002d01000005000000060101000000 ,
                        0x08000000fa0205000100000000000000040000002d0101005700000038050100 ,
                        0x2900600050020902500217024e0224024a02300243023b023b02430230024a02 ,
                        0x24024e0217025002090250020902500260004e0251004a024400430238003b02 ,
                        0x2d003002250024021e0017021a0009021900090219006000190051001a004400 ,
                        0x1e00380025002d002d00250038001e0044001a00510019006000190060001900 ,
                        0x09021a0017021e002402250030022d003b023800430244004a0251004e026000 ,
                        0x50026000500208000000fa0200000000000000000000040000002d0102000500 ,
                        0x0000060101000000040000002d01000008000000fa0200000a00000000000000 ,
                        0x040000002d01030007000000fc020100000000000000040000002d0104005600 ,
                        0x000025032900600050020902500217024e0224024a02300243023b023b024302 ,
                        0x30024a0224024e0217025002090250020902500260004e0251004a0244004302 ,
                        0x38003b022d003002250024021e0017021a000902190009021900600019005100 ,
                        0x1a0044001e00380025002d002d00250038001e0044001a005100190060001900 ,
                        0x6000190009021a0017021e002402250030022d003b023800430244004a025100 ,
                        0x4e026000500260005002040000002d010200040000002d01000004000000f001 ,
                        0x030007000000fc020000c0c0c0000000040000002d0103000500000006010100 ,
                        0x0000040000002d0101002a0000003805020005000d000a01a600d10034016d01 ,
                        0x3401a601a6000a01a600c300c2018901c201a601a601c2017b01c2012601a601 ,
                        0x0a017e010a01a601a6000a01a600d9001d01a6005001a600a601c300c2010400 ,
                        0x00002d01020005000000060101000000040000002d01000008000000fa020000 ,
                        0x0a00000000000000040000002d010500040000002d0104000e00000024030500 ,
                        0x0a01a600d10034016d013401a601a6000a01a6001e00000025030d00c300c201 ,
                        0x8901c201a601a601c2017b01c2012601a6010a017e010a01a601a6000a01a600 ,
                        0xd9001d01a6005001a600a601c300c201040000002d010200040000002d010000 ,
                        0x04000000f001050008000000fa0200000a00000000000000040000002d010500 ,
                        0x040000002d0104000a00000025030300a600a6019701a601c2017b010a000000 ,
                        0x250303008901c2019701a601970150010a00000025030300a600500197015001 ,
                        0xc20126010a00000025030300a6010a017b0134016d0134010800000025030200 ,
                        0xfb000a015f010a0108000000250302001201d8007501d800040000002d010200 ,
                        0x040000002d01000004000000f001050004000000f001030007000000fc020000 ,
                        0xffff00000000040000002d010300040000002d0101000e000000240305005001 ,
                        0x810150019201890192018901810150018101040000002d010200040000002d01 ,
                        0x000004000000f001030007000000fc0200009a9a9a000000040000002d010300 ,
                        0x040000002d0101000e0000002403050050016401500175018901750189016401 ,
                        0x50016401040000002d010200040000002d010000040000002701ffff04000000 ,
                        0xf00103000300000000000000000000000000000000000000000000004e414e49 ,
                        0x01000000ffffffff0e0000000000000001000000ffffffff400000004e060000 ,
                        0x4d09000006000000100000001800000000000000000000000000000000000000 ,
                        0x0000000002004f006c0065005000720065007300300030003100000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000180002010200000006000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000700000058140000 ,
                        0x0000000056006900730069006f0044006f00630075006d0065006e0074000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001c000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000012000000e5430000 ,
                        0x0000000056006900730069006f0049006e0066006f0072006d00610074006900 ,
                        0x6f006e0000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000220002010500000008000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000180000001c000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000001900000038000000 ,
                        0x00000000ffffffff0e0000000400000001000000ffffffff1000000000000000 ,
                        0x4e0600004d09000030140000010009000003180a00000600e60700000000e607 ,
                        0x000026060f00c20f574d464301000000000001002e770000000001000000a00f ,
                        0x000000000000a00f0000010000006c00000001000000010000003d0000003d00 ,
                        0x00000000000000000000a30500003e08000020454d4600000100a00f00006800 ,
                        0x000003000000000000000000000000000000560500000003000040010000b300 ,
                        0x000000000000000000000000000000e2040038bb0200460000002c0000002000 ,
                        0x0000454d462b014001001c000000100000000210c0db01000000600000006000 ,
                        0x0000460000004401000038010000454d462b3040020010000000040000000000 ,
                        0x803f1f4004000c000000000000001e4005000c000000000000001d4000001400 ,
                        0x000008000000020000005900000008400003e8000000dc0000000210c0db1700 ,
                        0x000000000000efbe1841959c6c42e4435042959c6c42a5eb5f42969c6c42959c ,
                        0x6c42a8eb5f42959c6c42e5435042959c6c42e4435042959c6c42e4435042959c ,
                        0x6c42e3435042959c6c42e4435042959c6c42efbe1841959c6c42ca3fb440a7eb ,
                        0x5f42ab701d40e4435042aa701d40e4435042a4701d40efbe1841a4701d40cf3f ,
                        0xb440a6701d40b1701d40c53fb440a8701d40ecbe1841a4701d40efbe1841a470 ,
                        0x1d40e443504298701d40a6eb5f42bc3fb440959c6c42e8be1841969c6c420001 ,
                        0x0303030303030101030303010103030301010303830314400080100000000400 ,
                        0x0000ffffffff2100000008000000620000000c00000001000000240000002400 ,
                        0x00000000803d00000000000000000000803d0000000000000000020000002700 ,
                        0x0000180000000100000000000000ffffff0000000000250000000c0000000100 ,
                        0x0000130000000c000000010000003b000000080000001b000000100000009900 ,
                        0x0000b3030000360000001000000042030000b303000058000000340000000000 ,
                        0x000000000000ffffffffffffffff060000008003b303b3038003b3034203b303 ,
                        0x4203b3034203b303420359000000240000000000000000000000ffffffffffff ,
                        0xffff02000000b3034203b303990058000000280000000000000000000000ffff ,
                        0xffffffffffff03000000b3035b00800328004203280059000000240000000000 ,
                        0x000000000000ffffffffffffffff020000004203280099002800580000002800 ,
                        0x00000000000000000000ffffffffffffffff030000005b00280028005b002800 ,
                        0x990059000000240000000000000000000000ffffffffffffffff020000002800 ,
                        0x99002800420358000000280000000000000000000000ffffffffffffffff0300 ,
                        0x0000280080035b00b3039900b3033d000000080000003c000000080000003e00 ,
                        0x00001800000002000000020000003c0000003c000000130000000c0000000100 ,
                        0x0000250000000c00000000000080240000002400000000008041000000000000 ,
                        0x000000008041000000000000000002000000460000006000000054000000454d ,
                        0x462b0840010240000000340000000210c0db00000000ce000000000000008fc2 ,
                        0x753f02000000020000000200000002000000000000000210c0db000000000000 ,
                        0x00ff1540000010000000040000000100000024000000240000000000803d0000 ,
                        0x0000000000000000803d0000000000000000020000005f000000380000000200 ,
                        0x000038000000000000003800000000000000000001000f000000000000000000 ,
                        0x0000000000000000000000000000250000000c00000002000000250000000c00 ,
                        0x0000050000803b000000080000001b0000001000000099000000b30300003600 ,
                        0x00001000000042030000b303000058000000340000000000000000000000ffff ,
                        0xffffffffffff060000008003b303b3038003b3034203b3034203b3034203b303 ,
                        0x420359000000240000000000000000000000ffffffffffffffff02000000b303 ,
                        0x4203b303990058000000280000000000000000000000ffffffffffffffff0300 ,
                        0x0000b3035b00800328004203280059000000240000000000000000000000ffff ,
                        0xffffffffffff0200000042032800990028005800000028000000000000000000 ,
                        0x0000ffffffffffffffff030000005b00280028005b0028009900590000002400 ,
                        0x00000000000000000000ffffffffffffffff0200000028009900280042035800 ,
                        0x0000280000000000000000000000ffffffffffffffff03000000280080035b00 ,
                        0xb3039900b3033d000000080000003c0000000800000040000000180000000100 ,
                        0x0000010000003d0000003d000000250000000c00000007000080250000000c00 ,
                        0x0000000000802400000024000000000080410000000000000000000080410000 ,
                        0x00000000000002000000280000000c0000000200000046000000dc000000d000 ,
                        0x0000454d462b08400003bc000000b00000000210c0db1200000000000000986f ,
                        0xd441da108541e214a741a073f6412ae71142a073f64185942842da108541986f ,
                        0xd441da10854135be9b4133eb3342d83d1d4233eb3342859428428594284233eb ,
                        0x33428192174233eb3342f21ceb4185942842986fd441c6b41842986fd4418594 ,
                        0x2842da108541986fd441da1085417de2ad41574fe441da1085417d900642da10 ,
                        0x85418594284235be9b4133eb3342000101018100010101010101010101010181 ,
                        0x0101144000801000000004000000c0c0c0ff280000000c000000010000002400 ,
                        0x0000240000000000803d00000000000000000000803d00000000000000000200 ,
                        0x000027000000180000000100000000000000c0c0c00000000000250000000c00 ,
                        0x000001000000130000000c00000001000000250000000c000000080000805b00 ,
                        0x00007000000010000000100000002d0000002d00000002000000120000000500 ,
                        0x00000d000000a9010b014f01ed014802ed01a3020b01a9010b013801d0027502 ,
                        0xd002a302a302d0025f02d002d701a302a9016302a901a3020b01a9010b015c01 ,
                        0xc9010b011b020b01a3023801d002250000000c00000007000080130000000c00 ,
                        0x000001000000250000000c000000000000802400000024000000000080410000 ,
                        0x00000000000000008041000000000000000002000000460000001c0100001001 ,
                        0x0000454d462b0840010240000000340000000210c0db00000000ce0000000000 ,
                        0x00008fc2753f02000000020000000200000002000000000000000210c0db0000 ,
                        0x0000000000ff08400003bc000000b00000000210c0db1200000000000000986f ,
                        0xd441da108541e214a741a073f6412ae71142a073f64185942842da108541986f ,
                        0xd441da10854135be9b4133eb3342d83d1d4233eb3342859428428594284233eb ,
                        0x33428192174233eb3342f21ceb4185942842986fd441c6b41842986fd4418594 ,
                        0x2842da108541986fd441da1085417de2ad41574fe441da1085417d900642da10 ,
                        0x85418594284235be9b4133eb3342000101018100010101010101010101010181 ,
                        0x01011540000010000000040000000100000024000000240000000000803d0000 ,
                        0x0000000000000000803d0000000000000000020000005f000000380000000200 ,
                        0x000038000000000000003800000000000000000001000f000000000000000000 ,
                        0x0000000000000000000000000000250000000c00000002000000250000000c00 ,
                        0x0000050000805600000030000000130000000f0000002c000000210000000500 ,
                        0x0000a9010b014f01ed014802ed01a3020b01a9010b0157000000500000000f00 ,
                        0x00000f0000002f0000002f0000000d0000003801d0027502d002a302a302d002 ,
                        0x5f02d002d701a302a9016302a901a3020b01a9010b015c01c9010b011b020b01 ,
                        0xa3023801d002250000000c00000007000080250000000c000000000000802400 ,
                        0x0000240000000000804100000000000000000000804100000000000000000200 ,
                        0x0000280000000c000000020000004600000008010000fc000000454d462b0840 ,
                        0x010240000000340000000210c0db00000000ce000000000000008fc2753f0200 ,
                        0x0000020000000200000002000000000000000210c0db00000000000000ff0840 ,
                        0x0003a80000009c0000000210c0db1000000000000000da108541859428422fe9 ,
                        0x22428594284233eb334281921742d83d1d4233eb33422fe92242859428422fe9 ,
                        0x22427d900642da1085417d9006422fe922427d90064233eb3342f21ceb418594 ,
                        0x2842986fd44181921742a073f6412ae71142a073f641ea18c941986fd441d43b ,
                        0x0c42986fd441333ddb4139c0ac41f84d154239c0ac4100010100010100010100 ,
                        0x0101000100011540000010000000040000000100000024000000240000000000 ,
                        0x803d00000000000000000000803d0000000000000000020000005f0000003800 ,
                        0x00000200000038000000000000003800000000000000000001000f0000000000 ,
                        0x000000000000000000000000000000000000250000000c000000020000002500 ,
                        0x00000c000000050000805a000000780000000f000000140000002f0000002f00 ,
                        0x0000060000001000000003000000030000000300000003000000020000000200 ,
                        0x00000b01a3028c02a302d0025f027502d0028c02a3028c021b020b011b028c02 ,
                        0x1b02d002d701a302a9015f02ed014802ed019301a9013102a901b7015a015602 ,
                        0x5a01250000000c00000007000080250000000c00000000000080240000002400 ,
                        0x0000000080410000000000000000000080410000000000000000020000002800 ,
                        0x00000c00000002000000460000003400000028000000454d462b0a4000802400 ,
                        0x00001800000000ffffff010000007d9006420ad71942d86ab54060b3d93f2800 ,
                        0x00000c0000000100000024000000240000000000803d00000000000000000000 ,
                        0x803d00000000000000000200000027000000180000000100000000000000ffff ,
                        0x000000000000250000000c00000001000000250000000c000000080000805600 ,
                        0x00003000000021000000260000002800000029000000050000001b0268021b02 ,
                        0x830275028302750268021b026802250000000c00000007000080250000000c00 ,
                        0x0000000000802400000024000000000080410000000000000000000080410000 ,
                        0x00000000000002000000460000003400000028000000454d462b0a4000802400 ,
                        0x0000180000009a9a9aff010000007d9006425d800e42d86ab54060b3d93f2800 ,
                        0x00000c0000000100000024000000240000000000803d00000000000000000000 ,
                        0x803d000000000000000002000000270000001800000001000000000000009a9a ,
                        0x9a0000000000250000000c00000001000000250000000c000000080000805600 ,
                        0x00003000000021000000230000002800000026000000050000001b023a021b02 ,
                        0x56027502560275023a021b023a02250000000c00000007000080250000000c00 ,
                        0x0000000000802400000024000000000080410000000000000000000080410000 ,
                        0x000000000000020000004c0000006400000002000000020000003c0000003c00 ,
                        0x000002000000020000003b0000003b0000002900aa0000000000000000000000 ,
                        0x803f00000000000000000000803f000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000220000000c000000ffffffff460000001c00 ,
                        0x000010000000454d462b024000000c000000000000000e000000140000000000 ,
                        0x000010000000140000000400000003010800050000000b020000000005000000 ,
                        0x0c025b003e00030000001e0007000000fc020000ffffff000000040000002d01 ,
                        0x0000040000000601010008000000fa02050000000000ffffff00040000002d01 ,
                        0x010039000000380501001a000a003b0034003b0037003b00390039003b003700 ,
                        0x3b0034003b0034003b0034003b000a003b000700390005003700030034000300 ,
                        0x340003000a00030007000300050005000300070003000a0003000a0003003400 ,
                        0x030037000500390007003b000a003b000a003b0008000000fa02000000000000 ,
                        0x00000000040000002d010200040000000601010007000000fc020000ffffff00 ,
                        0x0000040000002d01030008000000fa0200000100000000000000040000002d01 ,
                        0x040007000000fc020100000000000000040000002d0105003800000025031a00 ,
                        0x0a003b0034003b0037003b00390039003b0037003b0034003b0034003b003400 ,
                        0x3b000a003b000700390005003700030034000300340003000a00030007000300 ,
                        0x050005000300070003000a0003000a0003003400030037000500390007003b00 ,
                        0x0a003b000a003b00040000002d010200040000002d01030004000000f0010400 ,
                        0x04000000f001000007000000fc020000c0c0c0000000040000002d0100000400 ,
                        0x000006010100040000002d0101002a0000003805020005000d001b0011001500 ,
                        0x1f0025001f002a0011001b00110014002d0027002d002a002a002d0026002d00 ,
                        0x1d002a001b0026001b002a0011001b00110016001d001100220011002a001400 ,
                        0x2d00040000002d0102000400000006010100040000002d01030008000000fa02 ,
                        0x00000100000000000000040000002d010400040000002d0105000e0000002403 ,
                        0x05001b00110015001f0025001f002a0011001b0011001e00000025030d001400 ,
                        0x2d0027002d002a002a002d0026002d001d002a001b0026001b002a0011001b00 ,
                        0x110016001d001100220011002a0014002d00040000002d010200040000002d01 ,
                        0x030004000000f001040008000000fa0200000100000000000000040000002d01 ,
                        0x0400040000002d0105000a0000002503030011002a0029002a002d0026000a00 ,
                        0x00002503030027002d0029002a00290022000a00000025030300110022002900 ,
                        0x22002d001d000a000000250303002a001b0026001f0025001f00080000002503 ,
                        0x020019001b0023001b0008000000250302001b00160025001600040000002d01 ,
                        0x0200040000002d01030004000000f001040004000000f001000007000000fc02 ,
                        0x0000ffff00000000040000002d010000040000002d0101000e00000024030500 ,
                        0x2200270022002800270028002700270022002700040000002d01020004000000 ,
                        0x2d01030004000000f001000007000000fc0200009a9a9a000000040000002d01 ,
                        0x0000040000002d0101000e000000240305002200240022002500270025002700 ,
                        0x240022002400040000002d010200040000002d0103000c00000040092900aa00 ,
                        0x0000000000003b003b0002000200040000002701ffff03000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000700a5048300080410542c1018ae650d6d0043006f006e0076006500 ,
                        0x6e007a0069006f006e0069004e006f006d00690000006900000000000500bc04 ,
                        0x9a0108040a000000410064006d0069006e000000730075006d00690000006400 ,
                        0x00000000566973696f2028544d292044726177696e670d0a0000000000000b00 ,
                        0xe54300000084010014000000b4535101884200005d010000520000000000ffff ,
                        0xffffffffffff00fffffffffffffffffffffffffffeffffff0000000010835301 ,
                        0xdb6c046090d01200d0d01200767104600871af007d4dac0000000000ffffffff ,
                        0xffffffffed64eaf10318e9f2ffffff8300fff6f2fff3fbf00701f8f1008980ea ,
                        0xf116048016002000190080ff00c0c0c000e6e6e6ff00cdcdcd00b3b3b3ef009a ,
                        0x9a9a2100800066ff6666004d4d4d00337f3333001a1a1a001800000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000018000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x551aebf0fff2f012ebf01ce8f3a0faf4070febf0f6f1e6f54ae7f4043aebf044 ,
                        0xebf00100543501e8f380f2f1490f5b0ff2f1e0fb3101e6f568eae7f40febf054 ,
                        0xebf00200548518ebf03cdffc94019401e6f501ae99020000032e0405ebf006aa ,
                        0xebf007ebf008ebf009ebf00aaaebf00bebf00cebf00debf00eeae7f493e3f82e ,
                        0xc802550101f901c201e701020062010095fe1912011f16021f11e301022f0361 ,
                        0x00feeaf185c004eaf1fd990c13307a14ae47e1977a843fbe0340e6f5c7010222 ,
                        0x471310990257162215002b197a12b903661622130462022114052a9016069016 ,
                        0x07901608b4117a12930921a516eaf186c404eaf10380c3000f116f0648130724 ,
                        0x5f131b295046e0fbf03f161f8f1c191203901691041f16bd127a1206e6177a12 ,
                        0x07a2862c08862ce51819120a7c260b941e269d130cd02c0d30269d130e16e614 ,
                        0xf03f3f1287c804eb150f10f81e2f2a3f48133fb95c2e97ffcbe5e23f200000e0 ,
                        0x86f2f0ef410b0fba0a7914d02c010ad02c02d02c03d0276c2f7e254b36108f24 ,
                        0x5436d5159115099016c528e61b05a07d08e00c13663fc10319149115008f1991 ,
                        0x1567299115bd1fcf1bf83fc522aa19120b90160c90160d90160eaa90160f9016 ,
                        0x10901611901612929016139011eaf1a4d004eaf133340123e2f9fd6f02100439 ,
                        0x533b23a43d3a3712003d10371201605502a860553c1137120460550560550602 ,
                        0x6055073d109e4fb04ed54fe744eb01a7021060aa11371211605512aa60551360 ,
                        0x5514605515605516aa6055176055189f551990161a97621004dd531be4527a12 ,
                        0x1c92003b84d404eaf111304f48150001a3d80425050f1039519f696c3bde52a9 ,
                        0x00e452de5201c36702c3670304e452eaf1a7dc046c683951bb6aeaf169b7e004 ,
                        0xeaf1f70c13202066168c2a7f661608083c7f597f4127204021167e709d33a417 ,
                        0x9e1d7a1204917c1d05a41c0662089246d270a0232a917c09917c0a917c0ba417 ,
                        0xeaf1e9bde404eaf1b17169f03f8807138813384f6e3c19178d33e61455e0df34 ,
                        0x02738c03862c04862c4105862c852fa51ba752eaf1bee8040aeaf1510c140268 ,
                        0x65548f1c1bb510426c5d017d553b948c51eaf1bfec0402eaf1fa716f719f4715 ,
                        0x5a5f6c5f7e5f00905fa25fb04fc24fc45fe64ff84f0a592a1912149016151255 ,
                        0xc0f0045b92fc2758c301404c2693c9649732b93f80adc991a84c33d22ca3a881 ,
                        0xae7355c9a1e5c7a6739f00e9af819f939fa59fb79fc99fdb9fed9f20ff912015 ,
                        0x0baf1da47a1210b6a69d13a511a5bc1292a69d1313c5bc1454a4a69d1315e5bc ,
                        0x16a5bc17a5bc131820c8a59d131925c73b649115551b90161c90161d90161e12 ,
                        0x5519c8f4081a3bd03f88cf9ac6455b8c3f21c30001098f66c3017d3ad000df34 ,
                        0x9235ddc5a235ddc5b235ddc5a285a00730ad8a07308a44201507331608aa0db5 ,
                        0x090db50a0db50b4cb510ae1f1611620936a6073e136900c004c301faf19c0480 ,
                        0x09bf058366e7f4e594e3f85427578007102700cd30c9a255c5d7ae0035bf3022 ,
                        0xddfe1080068366191404564f684f007a4320df8c459023e9d643df55df361355 ,
                        0x0c0db50d0db50e0db50f3d10509fb30be6aeb4b62c33901634901625357c2636 ,
                        0x02e62213394161eaf1a16ac40492dfa4dfdffc95e3f838582757333f8bf62133 ,
                        0x9ef2f30ae7020be601385401e7d7afdbff723f843f20963fa8389df6aa74daec ,
                        0x05dae730de02db5508e45259bf0d4c1ff0da2d81b742f5616b0f3438ff4aff0e ,
                        0xec8812e8427069014d4894658ed16f1144b86201542911f36301703c1f017012 ,
                        0x314e1d480febf815f56168f264274128c7074000d111e92741d1092115123593 ,
                        0x12e8452bf88401bcc207b43b94013a9551023a9527b1f561858068333083ff30 ,
                        0x7a14ae47e17a8428b3c46b7634f102f3f80254301848058671184df884ac6dac ,
                        0x638bfff02608b4c73e2f0b3cb61b3fcd2f11e94b3f00e2ffebfffdff0f0fa53f ,
                        0xa57945278047202cf58edfd00fe20f6afa8b3087dc2f8094ffa6ff0ff0bbff20 ,
                        0x4f324f1e3a0600452b986134f5291f7e4f4d1f5f1fac3c048015bb2124913757 ,
                        0x95ded5bb21a649002115c31fd51fe71ff91f0b2f1d2f2f2f00a52f532f652f77 ,
                        0x2f892f9b2f574bbe05006d4f146f1a6210652c69aa4f781d34f10085519434f1 ,
                        0x656c5b446f075f195f2b5fe83d59b72127e59c3083321cc70d71ef60ac3fee6f ,
                        0x007627e3f42f8c3c6f3f4f020000763534507c0202507c035077f3551ee1251f ,
                        0x93301a6f00904fa24f701ffb0300910be1943497312011e900910099b7191ee1 ,
                        0x05926fa46fd0b66fc86f485a9f54329f18ff0091176b5f4876d48b0136e1e35a ,
                        0x3701408e5fa05fb25fc45fd65fa82fba26018ce86ffa63109fe573187f337f3c ,
                        0x7f504a3f5c3f6554b502a4d9d836f38300e5790e4329f03966b2c525e46386fd ,
                        0x0101bf35889891cf30c4c554affbaf608fa033a48adf963fa83fc0de72021704 ,
                        0x089d83e3d4fb03c0f1df03ef15ef27eb52cd8707f1a6b502c93588ad0463010d ,
                        0x2d8826b238811886c4c17161c4c200cbe1cbe1716222f16d01cbe1da42db4162 ,
                        0x22f10069b179c16d01000fc5b430db4200d1011169b10012dfb479c2b513fe00 ,
                        0x14c5b40015fe00160a94151afe001bfe04f652f97126b102036c04aa7f39c4c4 ,
                        0xc147cd991fbc4f10ce4fe04ff24f3d8c224d8f78af8aa5fe495f553048fde7b4 ,
                        0x810b4e6b645301a48f243508d6b22610e575e76f54df0b75020a91879f5aaf00 ,
                        0xab9f4e7f607f727f847792afa4afb6af52c8ae71daa91db627b3d5bcf9af400b ,
                        0xbf1dbf4e6642b351e6cae30bd4641030ff04ef16ef66fe7bcc3f06fff03f00ae ,
                        0xf90e4f50ff62ff2bbaba0fc7efd9ef00f00f021f141a5bc11ecf30cf08067161 ,
                        0x00160d5fcf781f8a1f9c1f1f8fb5cfc7cf08d9cfebcfc5f123565f14df4e66e6 ,
                        0x89013dcaf90a914e63099f1b9f2d96b32d01664cdf6c1f70df82dfaa1f4b6576 ,
                        0x3441024493f2df94ffa6ffc8ae613aef0aeca3553322c559ef6bef7def4e6600 ,
                        0xa8efd81fea1ffc1ff0ef02ff14ff26ff0038ff4aff2c3f6eff80ff962fa82fb6 ,
                        0xff00c8ffdafc7161967fa87fba7fcc7fde7f0078cf8acf9ccfaecf448fa300a8 ,
                        0x0f6b8e22c20f2dd50fe70fc01701ddf5e689015b0f1f211f331f299aa83738b4 ,
                        0x0e300238b203ddf5d31f653f773fc8aeddf0001d2f2f2f412f532f652fd24d8e ,
                        0x71792100edffffff2b6638b1396d610f114f234f00354f474f8c0f9e0f7d4f98 ,
                        0x4f01dfbc4f0025df37df49df447f8417829fa01f6d5f00addfbfdfd1dfe3df57 ,
                        0x5f695f7b5f2bef009f5f4fefc35fd55fe75f97ef7a2fea7f00fc7fb02fc22fdd ,
                        0x2fe62ff82f0a3f1c3f004e9f403f523fb88fca8f883f9a3fdbfb00c993c03f24 ,
                        0x6ffa94b79108adff3f636f00756f876f996f594f6b4fcf6f8f4fa14f00b34fc5 ,
                        0x4fd74fe94ffb4f0d5f1f5f315f00435fe57f879f999f8b5f2d8faf5f518f1063 ,
                        0x8f758ff75fb8900b0d6f1f6f4cc90051a15acd516f400dab81777084b52a0d00 ,
                        0x1273619fae6fc06fd26f5314d2853d7f0019df617f737f57df977fa97fbb7fcd ,
                        0x7f0070bf82bf94bfa6bf278f398fdcbfeebf0000cf818f938fa58fc7dfd9dfdb ,
                        0x8fed8f00c9ef119f239f359f19ff2bff6b9f83ef0095efa19fb39fc59fe5f151 ,
                        0xa1df9ff19f0003af15af27af39af4baf5daf6faf81af0093afa5afb7afc9afdb ,
                        0xafedafffaf5e0300ecf620bf32bf44bfab3164b5cd0267bf00bbdf5dff6fffaf ,
                        0xbf03efd3bf27ef39ef444bef1bcf0d2ecf40cf1f290eec3400312972cf84cf96 ,
                        0xcfa8cfbacfcccfdecfc4f0cf02dba611dfed2f35d40fc003c0c040df52df64df ,
                        0x76df88df9adf1c4f345db205620f421f541f661f00781ff9df0bef1defc01fd2 ,
                        0x1f53ef65ef0077efa33fb53fadefbfefa54fe3eff5ef0007fff54f075f3dff5f ,
                        0x4f714f73ff85ff0097ff07623121b2ffc4ffd6ffe8fffaff000c0f1e0f300f42 ,
                        0x0f540f660f780f8a0f209c0fae0fc00fd20fa0420e10311bbf08b24f121f51b6 ,
                        0x0e5bbf401f2e5f405f00c23f881f9a1fac1f0a4f1c4fe21fcf99090f2674935c ,
                        0x01a35ffc742dc10a8d003e7f572f692f7b2f8d2f9f2fb12fc32ffa6d6b337c64 ,
                        0x48fde7b481034e6b8a6f9c664633c786d296a3d100e17ff37f219615612f9d22 ,
                        0x8f348f468f20588f6a8f7c8f28719287009d8ead67028245476c58ee0fdd6fef ,
                        0x6f51b925c5006b51069f189f2a9f3c9f4e9fe95ffb5f200d6f1f6f316f918f55 ,
                        0x6f035044b48ffe816d460a8542a150288bb43fdf8d0341b6ffc6de556c88e42f ,
                        0x7fbf35d946153fb8bf393f0600a6b652421512e9f215e3f801dcf0000023166e ,
                        0x015546ebf0fff2f002ebf044e8f366faf40002f1f20e0fffffeaf1501f0de0fb ,
                        0xf6f1e6f568e7f404ebf0bd28ebf001005418ebf01050dffc59015901e6f50142 ,
                        0x0403e7f4f592e3f8a4060255464d26bf93c96432d93fab078cbfc7e3f1783cbe ,
                        0xbc06bfff460a8542a15028a4d13fcf067c09e3f850e0fbf03ffce60fe1fa01ff ,
                        0x03000420ae670203050aebf01febf0f2922e13043517eaf1c08504eaf1906ca5 ,
                        0x038306464cad02b93f6f17faac03c98018b95c2e97cbb3e5d29218701e7355b8 ,
                        0x11e5a0b616161fd81f84043205c84208488ca503c614d03f0d2f1f26011bf02b ,
                        0x3f0156140721160f4922a401d290058a6018e7f4fe20021004d1fd3006071757 ,
                        0x125d06011b60ff29730069006d0062ef006f006ca8202c0073ff00650067006e ,
                        0x0061baaa2065ae206c0075a82067aaac2269b62074b22072b22073aeb0222c00 ,
                        0x53d02061a4207080b820cf237620eaf1f121880188011c105e07e10dfe25e2f9 ,
                        0x83e3f83511a80002351147dd16eb8f01eaf120f3f0000081feeff20df0000007 ,
                        0xc0001b0001050180830b0f1d0f2f0fc8410f530f5d040c0505010186ffff8c00 ,
                        0x040fffff008cffff04f00fff0f8dff3f020ff08eff82910f97007f84ff850001 ,
                        0x0f849700ff83ff01f0847782707eb003860002070f83b102ff080777777bbb77 ,
                        0x004ec608788877c505c20170c703df0107847783ef05f084df0082070100c701 ,
                        0x84ff9d01c60003f070ef02b00182870003f0d4030e100d12b001853fff050f00 ,
                        0x000f31141011602c1140120310b001960f03f08b01f3030f84057e00010f91ff ,
                        0xffc001000054007200ff61007300630069007d6ef5f07200650020fdf0576500 ,
                        0x6c0b0061050070f5f07567fbf42c0500710075fbf25564fbf020250074fbf06c ,
                        0xfbf0517a3b0000052e017025006cf7f04961fdf03401202b000a0762fff254f4 ,
                        0xf15c0367370220f7f074f3f0a775006d0300560570030072ae05006d006f2b02 ,
                        0x66fbf0632a3f086c0500639f006c9f0002032a2c0372fbf065890070fbf08a05 ,
                        0x416f0500c4030a016801c001649f00552e050046fff663370263b704016eb106 ,
                        0x4c0be6058203e201ea059e010175f7f004010401f8f178059801040175503b12 ,
                        0x70cd047400e0930802d801709f008201ac0b2e118201fa0103000000009548eb ,
                        0xf0fff2f002ebf4ebf0443aebf004ebf002000aeff4110d1111eff42309e4f701 ,
                        0x06023b07eaf16968e7f40501283e025418ebf0a110dffc05010501390d03e7f4 ,
                        0x9bfae3f8ed060255464d26935fc96432c93faf0dd9c00edab608508109c0036c ,
                        0x010302dd053e0401041606010475f6f6f06020e8f3e03f05fe6af40203fa0605 ,
                        0x0412057594008a0e191a0602740b110b01351d01a97431114b1c140601024616 ,
                        0x628b010072140360168312eaf185fa4a0833a903307a14ae47cfe17a843f3c03 ,
                        0xd304a93f22f6f1024c036c018c103062149b12158642083da90301f2f04709fd ,
                        0x14b0bb131129e509ebf0f03f4d04a32891083e21ac00fef2f0fdf2f05121f851 ,
                        0x21350ba0110201603d43ff006f007000790072af006900675400744e1028ef00 ,
                        0x6300294e10320030aa9420334e104d8020637e206f4b00737820668622772172 ,
                        0x7a205d2e4e1054007586207480208520c4226480207f21c1257f2173bb00657e ,
                        0x20760061c2222e8ae2156c3e0406ebf0dd104000550068016801e2f900310031 ,
                        0x840f0303fa01eaeaf189e3f81ba90300010dae68010201700b0121ebf0f27f00 ,
                        0xc0801f0000409b12458a4a0854a904e6f58c3601d51365000a12623d30191202 ,
                        0x013012d2ab36033e32fa010a6b02030114c6364d028b420834a90ac800963f00 ,
                        0xa8318411b03de0359405f03ec205094f00b0399815e4318005f0358d364c4fa8 ,
                        0x3fa4704b3f3532863f3b050f0602725e41010a00c0feec4201f2430301c09b12 ,
                        0x9548ebf0fff2f002ebf4ebf0443aebf004ebf002000aeff4110d1103eff42309 ,
                        0xe4f70106023b07eaf16968e7f40501283e025418ebf04110dffc05010501390d ,
                        0x33059be3f8fdf9060255468fc7e3f1ff783cce3f4669341aff8d46a3c13f460b ,
                        0x85ff42a15028a43f460cbf0683c1603088c005946ec905783f508109c0036c01 ,
                        0x770302063e040104160601db0475f6f06020e8f3e03fab05fef40203fa060504 ,
                        0x1205297523000e191a0602740b110b016b20335012e33d1501743111aa4b1266 ,
                        0x6a12d63d150246179ad5998511b93d15036017b81e9f85eb51b89e1710eaf1a0 ,
                        0x404a086801ac00350d3d036c038642085d3da90305ffff460a01f2f0c2ebf03f ,
                        0xe6f50526e509ebf0f03f45014d03a391082f21ac00fef2f0a1fdf2f042214221 ,
                        0x350b855b1260ff3d43006f00700079bf007200690067540074bb002060006300 ,
                        0x29792032ab003085203379204d7120632e6f206f00736920667722682175726b ,
                        0x202e792054007577201574712020b5226471207021b225ee70217300656f2076 ,
                        0x006192b3222eebf0a7126c3e04fa0130023e025568016801e2f9fa01fa01840f ,
                        0x8a030305e7f489e3f82301ac00022b010feaf18a4a0854a904e6f52c6536d313 ,
                        0x02000a12622e3019129302013012843603f02230310aa66b0203019f364d028b ,
                        0x42083410a904c2066f3f813101863fa712bd31003305c93ecb05e23f8939f031 ,
                        0xb935800540c9356636254f813f494b3035325f3fdc3b054d3102007241010a00 ,
                        0x6bc0fec54201cb4301c0a7129548ebf0fff2f002ebf4ebf0443aebf004ebf002 ,
                        0x000aeff4110d1103eff42309e4f70106023b07eaf16968e7f40501283e025418 ,
                        0xebf04110dffc05010501390d33059be3f8fdf9060255468fc7e3f1ff783cce3f ,
                        0x460b8542efa15028c4b705a43f467f0c0683c1603088b705dd94c905783f5081 ,
                        0x09c003ee6c010302073e04010416b606010475f6f06020e8f3e0573f05fef402 ,
                        0x03fa0605041253057523000e191a0602740b11d60b0120335012e33d150174ac ,
                        0x31114b129a996b11d93d1502d446176a13b93d15036017b81e9f85eb51b89e17 ,
                        0x10eaf1a0404a086801ac00350d3d036c03864208bd3da903129a9a9a47090184 ,
                        0xf2f0ebf03fe6f50526e509ebf0f08b3f014d03a391082f21ac00fe42f2f0fdf2 ,
                        0xf042214221350b855b12ff603d43006f0070007f790072006900675400777400 ,
                        0x20600063002979205732003085203379204d71205d636f206f00736920667722 ,
                        0xea6821726b202e79205400752a772074712020b5226471207021dcb225702173 ,
                        0x00656f207600a561b3222eebf0a7126c3e04060aebf0303e025568016801e2f9 ,
                        0xf12128f121840f030305e7f489e3f82301aeac0002010feaf18a4a0854b0a904 ,
                        0xe6f56536d31302000a1262cc2e301912020130128436030034fa0130310a6b02 ,
                        0x03019f364d02858b420834a904c2066f3f81310100863fa712bd313305c93ecb ,
                        0x05e23f893900f031b9358005c9356636254f813f494be23035325f3f3b054d31 ,
                        0x0200725e41010a00c0fec54201cb430301c0a7125547ebf0fff2f009ebf044e8 ,
                        0xf346faf40007e9f2f2f1110d11eff4502309e0fbf6f1e6f5683e083cebf02f01 ,
                        0x005418ebf024360f440a5501ebf002ebf003ebf004ebf05505ebf006060408e7 ,
                        0xf49be3f8fd698800020055464d26bf93c96432c93fbf0dd974d00ec608507d09 ,
                        0xcf037e8d019b020a850401040a118c010185030a112b420298010e1227170208 ,
                        0x9c012c133c160209012c1350160912662212040b29166e1602168901db04758c ,
                        0x006020e8f3e03f4705fe6a8d011d120e1205281655012816033c16013c160350 ,
                        0x16a90150161d17056e16016e1603a68212057590008c19108d0106320a110c5e ,
                        0x019c0100848504eaf1c55b890255e2f93022222102027f804c000040fe1a8d01 ,
                        0x590d0a11230101100a112aebf01bf303680103011e2609122222941d128d11f2 ,
                        0xe9f2878904eaf15cde3923321cc771b220ac3ff8b12fc326302346b95c2e97af ,
                        0xcbe5e23f8d10e0f2f0ef9141dcff422beaf19c8d04eaf114feb90c40dcfd66b1 ,
                        0x16eaffc3bf40cf568dfe265b7dc52c35d33f3634b52c36aef40b00000f890272 ,
                        0x2213c0f5fe6932016f3301c0fe155a5122755720608188012d5821651f890103 ,
                        0x8d328d3b802e582104182285270482149c308c195c221d128a2222051e160102 ,
                        0x248d108c1927f6890106759c0060517073bf2d3852c100401c44f04bbf819001 ,
                        0x605821eaf1be910442eaf11939243b12fa2c3025859504faeaf1633923307a14 ,
                        0xae47cfe17a843f8303e304a93fc28c01022f231821a030891162080b00061723 ,
                        0x03980118131822f7124a181201182209c8464502a4990472eaf171392fe8f310 ,
                        0x04fef2f0d2e8f3f087432d253489011062ab00002c507beaf1c158206013013a ,
                        0xebf02e42803852b6304057e81723f032942786060600009d02f343010e563023 ,
                        0x815418561856f5090014528800b2436e119001c0446d12c852a9045a53c45605 ,
                        0x182208c4560884d247c85209bb336e1197309724a30aa108bcf343fdf2f03561 ,
                        0x0d510d577c8205484360072100232550df360033003560600000fdfe6c410201 ,
                        0x603d4300ff6f007000790072005769006750007493202877405d299320320030 ,
                        0x9160339320754d7d60637b606f00737560a966836274617277602e932054ab00 ,
                        0x758360747d6020c16264707d607c61be657c617300657b6047760061bf624c51 ,
                        0x37422c8504020901145e030171defd7255eaf11861fce6f564010200543c006e ,
                        0x2e756020006de560738360e061413e156430255c21387f4a743b1564aaa2250e ,
                        0x607365a36063876030a08b60166315358d1163704e4472633a75606ec9622000 ,
                        0x262e50b0612164756016634045347441677068712574e5607a7d60bc71651564 ,
                        0x7045152860734cbb606f7f62c665c461016e48749371de611663514501716370 ,
                        0x5553ee726d7760612784007f60828a030d31746104778f6783e6f5c912e3f8a0 ,
                        0x607322211eac8430229c31909c32a9758d11153221d38440422204e084704223 ,
                        0xed84272268016801090209250794a402261494450211411141800a110a820b81 ,
                        0x6e119c729d7122210035b42d28356073e95f30240d51fb0101179104008401d0 ,
                        0xfae9f2c0e8f3465d00505dff7c5338479c46130c0f74335a6d4795a2259d7163 ,
                        0x70668d14f03f67925e74745131422f00007452c1940637424b910815353e2163 ,
                        0x70613025b39318222b92dce8136761020075aa2060a1440a918d37c538513e51 ,
                        0x8c01c35820419242529c3036434b91404533e49f486a9076226f460a7b222211 ,
                        0xf0fe9681f015324b917045579f699f7b9520f76313058c018e4000600f8cf305 ,
                        0xb2479551450b84aa96b694010101484201769052ff840a11f081f37530479572 ,
                        0x55e5aa80034a402812faaf680110b9a405e5aa072061403c1239bf0a9110b945 ,
                        0x05025f74e83561ef97c843c391d2962c9280ebf06933a995f19547953992d0be ,
                        0x264140a9a3479546925c77dab540356100f6137060ed659291900198016d8fe1 ,
                        0xfa047045948d386073940170217021302240f085a222fd8515320a954042b62d ,
                        0x2885afa3b4fcabb82fc2a6a49400107f040001604000008733576005542f8274 ,
                        0xcf72fe7c714302032c5f3e5f2dd87c71081723b2cab608bd546c26001b890262 ,
                        0x170100727023c02ba25552c2c592a2257aa3bf386c01f5c59c320257600a5350 ,
                        0x0061776065ad60456c4472736861d68157600be77f4af973fe8211052b92f081 ,
                        0xf2634088f2f1c2c515358899dfabdff7c31d8afec309cbd7547960d1d1696031 ,
                        0x02546313128f248ffedc680111ed4045f098df3beff6c4a3c202600c5346bb60 ,
                        0x62005cef696054e2005ae1104f8f02e80a9111e96e9104900150cf00e1fa1535 ,
                        0x948daf8163708c012211221162880111f561f5618c0111a92d2801da60739001 ,
                        0x3565dab6a6a43d6f8204f13030436c31d681f200c0a236c09000a0f561600fb4 ,
                        0x7fc67bbda30a9193600d4da162749a7b6061c37e90fe35c8c842020cbeb22211 ,
                        0x020301b80804070539a2218004103643b5f1a2258160733565f0d0ffe2fff4fa ,
                        0xa3c3600c2500555fc570509f62707b62658360f1e06806660e6932036220050b ,
                        0x8082344570e884974159fffe4a40a225948d66216370a2119bfaaa2d28253a1e ,
                        0x3254243fac1f0297402fc255c0a14300a84302af40127932ac0142e611e41377 ,
                        0x33026f3b03bc7e36807208601152cf607002a16269f2764d01497109d665a3e7 ,
                        0x3608fe97c0327f3204482274a2eb3274a6310962d203a3e38102046f35e1166b ,
                        0x0de4e61005e61003fe2652e5057e35750860a22c620500597264539603050337 ,
                        0x4266f584820ec11c4b17404519599d159955a1327cc854f34ad9c600e86918dc ,
                        0x0e6a1c537af8b30218d080c2983100f9b376d38fb48c7172558c714e10995501 ,
                        0xfc85b2fb715bb151f10b814e14de5140ff4951f101390ec1974178f5837cc808 ,
                        0x2c917c40b841487ec6f8319d1504440039910f4b9b95044473b10f4b78f50444 ,
                        0x0072510f475548ebf0fff2f006ebf050e8f37544ebf00cebf005000aeff40211 ,
                        0x0d03eff42309e4f72301f6f1eaf155023e0404e7f468e7f405ebf0bd2cebf001 ,
                        0x005418ebf014d0dffc65016501e6f5014a0f009bfae3f8f9950255464d26939f ,
                        0xc96432c93fbb0fcd0db9b8de07faf48f03c003103e0102dd0b91040104169501 ,
                        0x0475f64d006020e8f3e03f05fe6a0012030616051012057523008a1a191a9502 ,
                        0x7417110b01411d0139743d11571f000202521f4a142d036c1f05feeaf1859104 ,
                        0xeaf1fd43b503307a14ae47e1977a843f8f0340e6f54d0102c8bb1300116a100f ,
                        0xa214251202091374cdf615eaf1a09504eaf17001e8b800e2f9bc1501ed120000 ,
                        0x86e43e04eaf1199400b8000ec0c005c02b2901f2f0ebf0d3136829f109deebf0 ,
                        0xf03f010168020203fb722debf0f200c0a21d4a0f2113ebf07474008a12a30b01 ,
                        0xbf9362000090fe91220456972f7415ab2b01ba2505c32fd1741010ad2f960008 ,
                        0xef2f74d034033f9122091b3f74d12f3feaf151a352087131b800fef2f0fdf2f0 ,
                        0xf084318431252bb8110201603dff43006f00700079005f72006900675c00745a ,
                        0x10df28006300295a1032005530c730335a104db33063b130976f0073ab3066b9 ,
                        0x32aa3172baad302e5a10540075b930740ab33020f73264b330b231f435b23177 ,
                        0x730065b130760061f532052e0a2567910870016d04414f14272aeaf1c9e3f811 ,
                        0x95025494014831544831bc12b4b9185c6f4361e6f58e84310100383e02061123 ,
                        0x0101542b12b24204ae4605ae4608ae46e109d6204f11a5208a12801f004b0040 ,
                        0xb3126c9504f6f1306a02815570013e4df6f14505900f86078974360c222101f5 ,
                        0x1000008ab918fd5ab50469341a8d46a3adb1cc0701001012001613661a9552d6 ,
                        0x231402013c12e420241434ad4265010aed120301b556bc12f18b1328da11b801 ,
                        0x8fc7e3f1a7783c9ede07875b330662c30023101012a453a91cd3513305df5576 ,
                        0x5389c1ee5f8f59e60d6f1f6c55053afeb504acd56ab55aadc62c7d5f1613cdcc ,
                        0xb561ec9c5f2566e28805326f5fa0655d510200725e94010a00c0fef96201ff63 ,
                        0x0301c0f44623016e4192410954010140e2f96e416e41285f3a57f6f107ebf095 ,
                        0x08ebf009ba5200ad4205010d82ebf00eebf05d513b23465fb80001f85c5f8d65 ,
                        0xe6f5460b8542a1075028944e60ee128e53b820656bcb9a99f471b92368d75d5b ,
                        0xad3f56abd56ac53fcc78fa5f41eb646cf37f33058d65ce0ca9d47c84ab51676f ,
                        0xd0b1146251856dcd73c4113fc3764f6b4a83e9bb6ce371ce6d000d8f1f8f318f ,
                        0x438ffe744505069fce7211b41e9f309f4294d9a2886a71b87f00ca767491d77f ,
                        0x909f52876e71599f6b9fc07d9fe09f51887271608fd400ef769fbbdd6eb7bb7c ,
                        0x8f35af99e1e1a2880b01db599566190c060f83c160c0a56fb7648f89b463f1e4 ,
                        0x51880611db59ba5c2e9727cbe5c2456f1619e8646f1f6c000501326f446f566f ,
                        0x36bfa2888271db5f09aebcaf0069d3dbafedac8671aeb900c3b5c2bfaf68f0a1 ,
                        0xe3bff0a562515d52fedc58ca6432994c26b3bb3f466ac193c9c473bc523fb81e ,
                        0x85eb51d889bc955321ea51888e75df5572c26412b071cffe18115c8fc2f5285c ,
                        0xe7008ecfa0c8f94155018671faf10954a84100e2f9867101c54e7f607f727f90 ,
                        0x7f545f00665dcd75c58fd788f971e68daf578271010cbe58e2d2ce5eb58fa9df ,
                        0xd98feb8f002d6e9566179fd0af3b9f7d6f668f788f008a8f9c8ad2695caf6eaf ,
                        0xd4ef92ac519f40b3af71ef83efeebfab9660cbc361ff0081c931bd8ffffd958d ,
                        0x6f9f6fb16fc36fc053aee66f75bf9a5ff388a5ad8944ff22914824a23f464fbf ,
                        0xa7d3e9743abd690c0a3fd7a3703d0ac789bcdfc2038fe2f1afb39fc59fd79f40 ,
                        0xaf51b700f90fce7f1d1f95efa9b6de659fdff869050c057b0c14748671405553 ,
                        0xb13b40000115a5a15f5c01000048ebf008eff0df27002b0002e9f2c364d78c1f ,
                        0x06ebf0ff0c000000ff92f0f690a128c13ef7000001fdf400005a3fbfca429068 ,
                        0xdb31e2f90cfae7f446ebf02c6751011def0d00006ceff0d60023feebf0a4fc41 ,
                        0x01890e00bb00d3ebf042000aebf03cfffe41015c0f0000fbfe61063482400157 ,
                        0x1000f2fdf240dcfff1f20c6851015559810023fef0d6f1f2eca100f37c124400 ,
                        0xa906cc6951011fa814000028a902d10fe30ffd47ebf0ac6a5101d016ef000051 ,
                        0x08aa05dc6c517f01211f00006305ce05010000001c0000001d00000000000000 ,
                        0x00554ae9f21cebf01debf01eebf0551febf020ebf021ebf022ebf05523ebf024 ,
                        0xebf025ebf026ebf05527ebf028ebf029ebf02aebf0152bebf02cebf02ddcf000 ,
                        0x00a574ebf034eff4eaf112e7f401ff58cd6ba1a289e33ffff8930075c22ced3f ,
                        0xd2e3f8021f04eaf10bebf0d4747f4d015d2500003eebf07d432d02b48340019b ,
                        0x3700022201410103dcf414000000160000000000000000000000000000000100 ,
                        0x00000000000040000000146651019d250000490000005200000000000548ebf0 ,
                        0x3cdcff040f160f280e1400000016000000000000000000000000000000010000 ,
                        0x00000000004500000084665101182600000d0000005200000000009590ebf028 ,
                        0xebf05cdcffe1fa01eaebf004e7f41eebf0bc6651bf0184240000c8ebf053fb00 ,
                        0xc9ebf0248a4d014cef25000011ebf040003fbeebf0dc655101e63d002eaeebf0 ,
                        0x500044ebf04c2700250126500569051400000000000000000000000000000000 ,
                        0x0000000000000000000000000000001800000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000534ebf028dcff040f160c5515 ,
                        0xebf0fff2f006ebf060e8f3fcfaf4f1f200008320002c06e0fbf0bfdcffe0fbf2 ,
                        0xf109034c0f530000f6f1e6f546e7f403ebf09d44ebf00100547901eaf1028042 ,
                        0x06900d0507a407e0fb7501e6f5686e810600002085025418ebf0a108dffcd801 ,
                        0xd801e6f501e7f492fae3f8a47202554658cd6bffa1a289e33f46f893ff0075c2 ,
                        0x2ced3f4673bfc6e3f1783cbe3316bfff46408742a15028a4513f46160715dffc ,
                        0x50130c3f5d1ffee1fa01ff0300043c00551072010507111febf0f2a513a904ac ,
                        0x17eaf19b0418911c13408ce6f5dd1df03fef1de4186f1901e3cf03a612fb0207 ,
                        0x110104169685010475d80060dc01e9f205d1fea612c4022c2205322205754875 ,
                        0x003c29eaf1652925ebf01cdd075104901d0418e6f583e3f80a1c13aa071147e5 ,
                        0xf6c9f808868502500a7a251d0414118d1f033f153f273fa0393f4b3fe5f6c421 ,
                        0x07110b760f0780021af2f1b921900f450f6d33c80d0508ebf01201e00414dffc ,
                        0xd231d231ff0920d80175018925ca11e6f545d7142313ffd33f4050ad4a24f1ff ,
                        0x67e33f40b85c2e9797cbe5e2394ed24b47132a0f02e9f2840418f1211f108d2d ,
                        0x1616fb05392d7e4fe8f31004feaf067e1fe3009cc10821454d46dcfd66ffb116 ,
                        0xeac3bf40cf561f8dfe267dc5f8452f40035351b5f8465d4ae7f4858628231c13 ,
                        0xff307a14ae47e17a845acb43013a44b23ff70c67d504808921e301e004895f87 ,
                        0x2bce21e6f5a0167202540cebf01ebe540712b01120b011fb02dc01dc01c40221 ,
                        0xe55489224122f254f532455145516102240c64528232251964e70226266409ea ,
                        0xf199273364b92100284064be32290c4d64bd5100b4f8086b41ba503d2443f03f ,
                        0xc0475d61c405696461071582c0476d7204d8017f5fea0ea25f00130038b75389 ,
                        0x212af4640712026520fb020f65c4021c658922b6041815730354e8c045c041c0 ,
                        0x410711c0418920051089286e86240711b56f8c2ed86f08bf53ba50d8012bb174 ,
                        0xed221201d63382ef2070cf34b16fc36f927f784854140711b47aaa04182fb753 ,
                        0xf04f3c8f1902d800c04700665e047a2f8c2f0adffca1e3f836a64a3e76c388f0 ,
                        0xf000e77f04177221e70186814361e004a61140925d86890149a92f1d12d80148 ,
                        0xa42a000715b9247501559bfb05b9248681559b80c405b924d231559743613162 ,
                        0xebf054fcdd03581f00004d2693c9cf6432d93fe9959e3fffff007755d231a835 ,
                        0x071175011201b921d80100e131e000e13df57da541e6f5e131ba505f53007400 ,
                        0x61a8607072a0016e70a07221a541071532216caf7120413b5294f70555917d04 ,
                        0x868103958a00f2f115a9a03f269d0319d00fe20ff40f000616154fef484d4c41 ,
                        0x4f534f02193951900715455f575fe4f76c0414f6f1300285025526b1e17df6f1 ,
                        0x6109014df5350989b90c81478a041868a3dc1fe7f4018b3bb85dc578b66ecfc4 ,
                        0x0588ce78b50076c786855dcec4cff5355dcfe1fa5591207501d2316f117d04bd ,
                        0x5108c8ae6f3100e3afaa3d2cd517b5094913b5778f898f20dffc52bf64bfeddf ,
                        0x76b5c208e79bbf0072527625da7fe38ffe7f1089a54183a600e8b11bb1f0bf02 ,
                        0xcf14cf26cf38cf4acffe08d69b4ea7d3e974baf8f6d776cf09d5a9d56ab55a23 ,
                        0xada608e7a1cffeebca38ffd4cf7f3f2010080402d10dff20ffcf06ffe3b67501 ,
                        0xf1215c96e780050069e9f129c0ef5f0382312ab13161b921146f31bd510db0c0 ,
                        0x0e66640e95d8ef58bf22efef0cdb8dc737119e16fffeb2c6030281402010d011 ,
                        0x3f7a0faacf8bbbb2cb445e01d6cd02a603ce97cfa0ff9d0f8406b4e5161f027c ,
                        0x02bed8088233680f511c89012ab570411f7e1ec695c00e31994cea90f1c4d808 ,
                        0x3e627cff3f4023121f89442291c86cf8be3509d57f148bc562b158cc63ff4056 ,
                        0x6451ff272d5e01300125ffb6fc1f803203be197f257e2c5e016b43b1c72c7e7b ,
                        0xb072b9bc3f40c9cf20475c2ecf6cf83c0509d5e3d7210397cbd32de8b1aed134 ,
                        0x0134d1d9f400ec61b24934019125fcff0e0f200f3e0f00e4ef620d2715fb0f7f ,
                        0xcef30faf0f80fd00271fd6cdc80f051fc31fd519351ff31f001148641dcc2326 ,
                        0x208b4f8f1d87ffb11ff8abff11fde51dcb66b3d96cff36ab3f4079bd5eaf07d7 ,
                        0xebc5052f6b1fa81f3b2d720f80ea46662dab3f23d6aed1b4afc6afff00b4e15f ,
                        0xdfefaf01bf97dfa9dfbbdfcddfe62afad63f4a278d53ae3f409d870120482492 ,
                        0x0847966482001aefe6bf99efabef423fcfef86373eb400973f0fdfbb3fcd3c94 ,
                        0x604a7fb80d8e6e00124f397ea07f01df437f5294d9a1ce5f02e05fff6411fd5f ,
                        0x096f1b6f2d6f3f6f00516f636f756dd6368e6fa06fb26fc46f00d66fe86ffa6f ,
                        0x0c7f1e7f307fee7f547f00667f787f9d9f9c7fae7fc07fd27ff79f0047eb0100 ,
                        0x000017000000170000000000000000554ee9f217ebf01cebf01debf0551eebf0 ,
                        0x1febf020ebf021ebf05522ebf023ebf024ebf025ebf05526ebf027ebf028ebf0 ,
                        0x29ebf0552aebf02bebf02cebf02ddcf00000a574ebf034eff4eaf113e7f401ff ,
                        0x58cd6ba1a289e33ffff8930075c22ced3fd2e3f8021f04eaf10bebf054557f41 ,
                        0x01ab2d000041ebf07d432d02d4874001ec37000022013a010400dcf414000000 ,
                        0x160000000000000000000000000000000100000000000000400000007c6f5101 ,
                        0xee2d00004a0000005200000000000548ebf03cdcff040f160f280e1400000016 ,
                        0x000000000000000000000000000000010000000000000045000000ec6f51016a ,
                        0x2e00000d000000520000000000140000005c0000000000000000000000010000 ,
                        0x0004000000000000001500000024705101fa260000a0060000d300c90000005c ,
                        0x8b4d019a2d00001100000040003f000000446f5101382e00002e000000500044 ,
                        0x000000b46f5101772e00002e000000500000000000000000005531ebf0fff2f0 ,
                        0x03ebf028e8f3151cebf00cdcff00f6f1f6f1e6f55501ebf002e7f44f22084426 ,
                        0x0261554101e8f3f2f1550fffff4e05906a09e0fb2901e6f56822080101021700 ,
                        0x5418ebf0047d0f8b0aeaf1b5bce3f82b160255022d01105caf0d7c0e1f0608e9 ,
                        0xf2c98908451e2602512d012815e9f206ebf005052a05482a043d0f2800580f69 ,
                        0x1380f6f1771db10f900fa20fb40fe1fa9b0ae3f845160251fa0f082f1a2f2c2f ,
                        0x0100fd66ebf0650100004100ff7200690061006c00df200055006ef9f06300df ,
                        0x6f00640065fff04d006953dcffdffcff3603e93f2f05f7013f603401020b0604 ,
                        0x05025201042401fd66ebf0050100005300ff79006d0062006f00a16cdcff120f ,
                        0x240f360f80eaf105ff050102010706020501073701fd66ebf00501000057007f ,
                        0x69006e00670064f7f4a173dcff180f2a0fe2f980eaf10500370afd66ebf04501 ,
                        0x000041007f7200690061006cdcffdc100f220f00877ae9f28008dee7f4ff0100 ,
                        0x404400ff022f0b0604025201042401fd66ebf04701000053007769006df5f075 ,
                        0x006edcffbc120fddfedf7b0061ebf080fd08e7f4ff01012000007f2820020b06 ,
                        0x0402520101042401fd66ebf0470100005000ff4d0069006e006700c54cf9f055 ,
                        0xdcff160fe1fadf7bdb0061ebf08008e7f4ff01ff012000002820020b17060402 ,
                        0x5201042401fd66ebf0470100004d00ff5300200050004700ff6f007400680069 ,
                        0x00f163dcff1a0fe5f6df7b0061f6ebf08008e7f4ff010120ff00002820020b06 ,
                        0x0405025201042401fd66ebf04701000044007f6f00740075006ddcff7c100f22 ,
                        0x0f00df7b0061ebf0fb8008e7f4ff01012000ff002820020b0604020252010424 ,
                        0x01fd66ebf0450100005300f779006cedf061006500f16edcff140fdffc870600 ,
                        0x04dee2f99f000020eaf1010aff0502050306030303002401ed66ebf04501f0f0 ,
                        0x007300ff7400720061006e00ff670065006c006f005520f5f064030073f7f061 ,
                        0xdcff1ee1fa40600080eaf13902e3f8070308062408fd66ebf0470100005600ff ,
                        0x720069006e006400b161dcff120fddfedf7bfff280fd08e7f4ff01012000007f ,
                        0x2820020b060402520101042401fd66ebf0450100005300ff6800720075007400 ,
                        0x3169dcff120f240f0004260febf0070200053c08fd66ebf0450100004d005f61 ,
                        0x006e0067f7f06cdcff24120f240f80250fe8f3043e08fd66ebf0450100005400 ,
                        0x7f75006e00670061dcff48100f220febf040260fe9f2043e08fd66ebf0470100 ,
                        0x005300df65006e0064f9f07900b161dcff140fdffcdf7b010280fd08e7f4ff01 ,
                        0x012000007f2820020b060402520101042401fd66ebf04501000052001d61f7f0 ,
                        0x760069dcff100f220f12ebf002260f3502053c08fd66ebf04701000044007f68 ,
                        0x0065006e0075dcff7c100f220f00df7b0061ebf0fb8008e7f4ff01012000ff00 ,
                        0x2820020b060402025201042401fd66ebf0450100004c001f6100740068f7f0ff ,
                        0xff110fce230f000010260febf0020001043c08fd66ebf0450100004700df6100 ,
                        0x750074f7f06d009169dcff140fddfe20260febf0020300053c08fd66ebf04701 ,
                        0x00004300ff6f00720064006900ff610020004e006500b177dcff1a0fe5f6df7b ,
                        0xfff000f7008008e7f4ff010120ff00002820020b060405025201042401fd66eb ,
                        0xf0470100004d00ff53002000460061001f7200730069dcff160fe1fa7bdf7bfd ,
                        0xf000008008e7f4ffff010120000028205f020b0604025201042401ed66ebf047 ,
                        0x01f0f00075001f6c0069006ddcff100f220fdf00df7b0061ebf08008fee7f4ff ,
                        0x010120000028bf20020b060402520104002401fd66ebf0450100005400ff6900 ,
                        0x6d0065007300d720004efbf077fff052009d6ff9f061006edcffddfe87ed7ae9 ,
                        0xf28008e7f4ff0100fd404400ff02020603051f04050203042401180000001002 ,
                        0x0000000000000000000000000000000000001800000000000000000000000000 ,
                        0x000000000000000000000000d7000000448b4901be2f0000450000004200d700 ,
                        0x0000cc8a4901033000002e0000004200d7000000548a49013130000025000000 ,
                        0x4200d7000000dc89490156300000350000004200d7000000648949018b300000 ,
                        0x390000004200d7000000ec884901c43000003d0000004200d700000074884901 ,
                        0x01310000430000004200d7000000fc87490144310000390000004200d7000000 ,
                        0x848749017d310000370000004200d70000000c874901b43100003d0000004200 ,
                        0xd700000094864901f1310000380000004200d70000001c864901293200002700 ,
                        0x00004200d7000000a485490150320000220000004200d70000002c8549017232 ,
                        0x0000230000004200d7000000b484490195320000390000004200d70000003c84 ,
                        0x4901ce320000220000004200d7000000c4834901f0320000390000004200d700 ,
                        0x00004c83490129330000260000004200d7000000d48249014f33000027000000 ,
                        0x4200d70000005c82490176330000430000004200d7000000e4814901b9330000 ,
                        0x3e0000004200d70000006c814901f7330000380000004200d7000000f4804901 ,
                        0x2f340000470000004200000000000000000000000000f528ebf001ebf046006f ,
                        0x007f72006d00610074f7f07d20fbf06f00760069fbf03765006efff200000100 ,
                        0x000047007500690064006100000001000000470075006900640065000000f528 ,
                        0xebf002ebf052006500dd74f7f020002dfdf04300a76f006e0700f8f174050072 ,
                        0x06f7f00000f520ebf002ebf04e0065007d74edf043006f006e010077650063f9 ,
                        0xf06f0072dcf0f51aebf001ebf043006f00dd6ef9f0650074fff06f000d72fdf0 ,
                        0x0000f518ebf001ebf043006f007d6ef9f06500630074f7f00172dcf0f516ebf0 ,
                        0x01ebf056006900fd73f7f06f00200039000130dcf0f526ebf001ebf052006500 ,
                        0xdd74f7f020002dfdf04300f76f006ef9f0720061000573f9f06fdcf0f51eebf0 ,
                        0x01ebf04e006500ff7400200043006f007d6ef9f07200610073f9f0030000f516 ,
                        0xebf001ebf056006900fd73f7f06f00200037000130dcf0f522ebf001ebf05200 ,
                        0x6500dd74f7f020002dfdf04e00ff6f0072006d0061000d6cf7f00000f51aebf0 ,
                        0x01ebf04e006500f7740020f5f06f0072001f6d0061006cdcf0f516ebf001ebf0 ,
                        0x560069007d73f7f06f002000300100030000f52eebf001ebf052006500dd74f7 ,
                        0xf020002dfdf04f005f6d00620072f7f0670d0057690061f9f075090061dcf0f5 ,
                        0x1aebf001ebf04e006500ff74002000530068007f610064006f0077dcf0f516eb ,
                        0xf001ebf056006900fd73f7f06f00200030000132dcf0f52eebf001ebf0520065 ,
                        0x00dd74f7f020002dfdf04c00d769006ef7f061fdf07300d56ff9f07405006cf7 ,
                        0xf00000f520ebf001ebf04e006500fd74edf0540068006900156eedf06c010265 ,
                        0xdcf0010000004e00650072006f000000f51aebf001ebf042006c00ff61006300 ,
                        0x6b00200017660069f7f06cdcf0f516ebf001ebf050006100df670069006ef7f0 ,
                        0x20000131dcf00000000000000000f52cebf001ebf04c007500ff6f0067006800 ,
                        0x6900f7200064fff269006e005f74006500720f007315000165dcf0f52aebf001 ,
                        0xebf050006f00ff69006e0074007300d520f7f066010049fbf2650035720f0073 ,
                        0xfdf00000f518ebf002ebf0530074005f61006d0070f9f06ef7f00165dcf0f514 ,
                        0xebf002ebf0500072007f69006e00740065f7f00300000200000053006f006c00 ,
                        0x530048000000f518ebf002ebf041006e00ff74006900530063001f61006c0065 ,
                        0xdcf0f514ebf002ebf048006100ff73005400650078000174dcf0f51aebf002eb ,
                        0xf053006800df6f00770042f9f072003764006501000000f528ebf002ebf07600 ,
                        0x69007f73004c00650067fdf0ff6e00640053007900ff6d0062006f006c000749 ,
                        0x0044dcf0f51aebf002ebf053006800ff6100700065004300356cf9f073050000 ,
                        0x00f518ebf002ebf053006800ff61007000650054000d79fbf20000f51eebf002 ,
                        0xebf053007500fd62f5f0680061007000df650054007901020000f516ebf002eb ,
                        0xf04c006100ff73007400540065000d78fbf00000f516ebf002ebf054006500ff ,
                        0x7800740053007900076e0063dcf0f51aebf002ebf0760069007f730056006500 ,
                        0x72f9f01f69006f006edcf0f522ebf002ebf0760069007f73004c00650067fdf0 ,
                        0xff6e0064005300680037610070fdf00000020000005400650078007400000002 ,
                        0x00000052006f0077005f00310000000200000052006f0077005f0032000000f5 ,
                        0x12ebf002ebf0530068007f610064006f0077dcf018000000ac04000000000000 ,
                        0x00000000000000002e0000002e0000000000000033000000ac8740014a360000 ,
                        0x04000000410033000000b48740014e36000004000000410033000000d4974e01 ,
                        0x5236000028000000470033000000348c4d017a36000010000000450033000000 ,
                        0x4c8c4d018a3600001000000045003300000004984e019a360000270000004700 ,
                        0x330000005c204d01c136000021000000470033000000a45a5401e23600001c00 ,
                        0x0000470033000000c45a5401fe3600001a000000470033000000e45a54011837 ,
                        0x00001900000047003300000034984e0131370000270000004700330000008420 ,
                        0x4d015837000022000000470033000000045b54017a3700001900000047003300 ,
                        0x0000ac204d019337000025000000470033000000245b5401b83700001d000000 ,
                        0x470033000000445b5401d537000019000000470033000000c4795101ee370000 ,
                        0x2d000000470033000000645b54011b3800001e000000470033000000845b5401 ,
                        0x3938000019000000470033000000fc795101523800002d000000470033000000 ,
                        0xd4204d017f3800001f000000470033000000648c4d019e3800000e0000004500 ,
                        0x33000000a45b5401ac3800001d000000470033000000c45b5401c93800001900 ,
                        0x0000470033000000bc874001e238000004000000410033000000c4874001e638 ,
                        0x00000400000041003300000064984e01ea3800002d0000004700330000009498 ,
                        0x4e011739000029000000470033000000e45b5401403900001a00000047003300 ,
                        0x00007c8c4d015a39000018000000470033000000948c4d017239000010000000 ,
                        0x450033000000045c5401823900001c000000470033000000ac8c4d019e390000 ,
                        0x18000000470033000000245c5401b63900001d000000470033000000c4984e01 ,
                        0xd33900002d000000470033000000445c5401003a00001d000000470033000000 ,
                        0x645c54011d3a00001a000000470033000000fc204d01373a00001f0000004700 ,
                        0x33000000845c5401563a00001a000000470033000000a45c5401703a00001a00 ,
                        0x0000470033000000c45c54018a3a00001d00000047003300000024214d01a73a ,
                        0x000026000000470033000000c48c4d01cd3a00000e000000450033000000dc8c ,
                        0x4d01db3a000010000000450033000000f48c4d01eb3a00001000000045003300 ,
                        0x00000c8d4d01fb3a000015000000470000000000010000000200000003000000 ,
                        0x0400000005000000060000000700000008000000090000000a0000000b000000 ,
                        0x0c0000000d0000000e0000000f00000010000000110000001200000013000000 ,
                        0x1400000015000000160000001700000018000000190000001a0000001b000000 ,
                        0x1c0000001d0000001e0000001f00000020000000210000002200000023000000 ,
                        0x2400000025000000260000002700000028000000290000002a0000002b000000 ,
                        0x2c0000002d00000000000000555ae9f203ebf004ebf005ebf05506ebf007ebf0 ,
                        0x08ebf009ebf0550aebf00bebf00cebf00debf0550eebf00febf010ebf011ebf0 ,
                        0x5512ebf013ebf014ebf015ebf01516ebf01aebf01bdcf0000000000000040000 ,
                        0x00a574ebf034eff4eaf116ebf002feebf00158cd6ba1a289ffe33ff8930075c2 ,
                        0x2c43ed3fe3f802010201eaf10bebf05f048644012816004bebf07d432d022496 ,
                        0x4201731600150aebf041e9f201dcf40000020000000500000006000000000025 ,
                        0x74ebf034eff4eaf103dcffddfee9021f04eaf10bebf0248d4dbf01cb3f00000e ,
                        0xebf0413e2d0204884001d9370022010541e9f2011004140000004a0100000000 ,
                        0x000000000000000000000f000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000400000006c7a51017d3f00004e00000052000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000040000000a47a5101db3f0000370000005200 ,
                        0x000000000548ebf03cdcff040f160f280e0548ebf03cdcff040f160f280e1400 ,
                        0x00004a0100000000000000000000000000000f00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000045000000147b510140410000 ,
                        0x0d00000052000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000450000004c7b51014d41 ,
                        0x00000d0000005200000000007bc402ebf0010000b8f3f04ee0fb010003ebf00a ,
                        0x0504ebf07d09ebf07701010022dcffe0340f460f580ff9fee4f7832000f12c86 ,
                        0x01680fe8f328a623f6ff1d93bee34028b497e3d01eab02a706a70e0871af5c79 ,
                        0x0aebf060194df7f614e7f47d16ebf0f4ef430180ebf0dd54ebf0420017ebf0ec ,
                        0x53575101d4ebf020ebf0500b125f24545101f81516fffb016b7b51f7f660940f ,
                        0x001aebf0f55c23101cf3f0f10b0000f7d2003cfb01894d010daf0d00000a0712 ,
                        0x43ebf0e4af914201176f1006ebf040fb001debf0a465510157dd260413520024 ,
                        0xebf0646e975101ab93101c1912ad1f214aebf09ca110cb9310181329ebf045d4 ,
                        0xa110ef93106110991027ebf0ff0c6f5101a92e0000f568191231ebf094775101 ,
                        0x5f152f0000a9ebf0d2ae1ffb00d803117951017634770000d0f3f0500032ebf0 ,
                        0xcd8c1f20103bf00016005400fd3febf0347a510112405700002a272244ebf0dc ,
                        0x4320035a41482561210000000000000000000000000000000000000000000000 ,
                        0x00000000feff0000040002000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000038000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x00000000feff0000050102000000000000000000000000000000000001000000 ,
                        0xe0859ff2f94f6810ab9108002b27b3d930000000d80000000a00000001000000 ,
                        0x580000000200000060000000030000006c000000040000007800000005000000 ,
                        0x840000000600000090000000070000009c00000008000000a800000012000000 ,
                        0xb40000000d000000cc00000002000000e40400001e0000000400000000000000 ,
                        0x1e00000004000000000000001e00000004000000787878001e00000004000000 ,
                        0x000000001e00000004000000000000001e00000004000000000000001e000000 ,
                        0x04000000787878001e000000100000004d6963726f736f667420566973696f00 ,
                        0x4000000070e47f23e154cb010000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000feff0000050102000000000000000000000000000000000002000000 ,
                        0x02d5cdd59c2e1b10939708002b2cf9ae4400000005d5cdd59c2e1b1093970800 ,
                        0x2b2cf9ae0500530075006d006d0061007200790049006e0066006f0072006d00 ,
                        0x6100740069006f006e0000000000000000000000000000000000000000000000 ,
                        0x00000000280002010700000009000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000001a00000008010000 ,
                        0x00000000050044006f00630075006d0065006e007400530075006d006d006100 ,
                        0x7200790049006e0066006f0072006d006100740069006f006e00000000000000 ,
                        0x0000000038000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000001f000000d8010000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000058010000140100000b00000001000000600000000200000068000000 ,
                        0x0e000000740000000f00000080000000170000008c0000000b00000094000000 ,
                        0x100000009c00000013000000a400000016000000ac0000000d000000b4000000 ,
                        0x0c000000dc00000002000000e40400001e00000004000000000000001e000000 ,
                        0x04000000000000001e0000000400000000000000030000000f270b000b000000 ,
                        0x000000000b000000000000000b000000000000000b000000000000001e100000 ,
                        0x020000000c000000506167696e61203100004c000c0000005374616d70616e74 ,
                        0x65004c000c100000040000001e00000008000000506167696e65000003000000 ,
                        0x010000001e000000080000004d61737465720000030000000100000080000000 ,
                        0x0400000000000000280000000100000060000000020000006800000003000000 ,
                        0x7400000002000000020000000e0000005f5049445f4c494e4b42415345000300 ,
                        0x0000150000005f565049445f414c5445524e4154454e414d4553000002000000 ,
                        0xe40400004100000002000000000000001e000000040000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Class ="Visio.Drawing.11"
                    OLEClass ="Microsoft Visio Drawing"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7650
                    Top =105
                    Width =576
                    Height =576
                    FontWeight =400
                    TabIndex =1
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3975
            BackColor =12311007
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =85
                    ColumnCount =2
                    Left =1680
                    Top =630
                    Width =6780
                    Height =1410
                    Name ="TopicList"
                    RowSourceType ="Table/Query"
                    RowSource ="qryTopicList"
                    ColumnWidths ="0;5760"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Select a Sample and click the Display button."

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    Left =1680
                    Top =2100
                    Width =6780
                    Height =1290
                    ColumnWidth =1500
                    TabIndex =1
                    Name ="Summary"
                    ControlSource ="memDescription"
                    FontName ="Tahoma"
                    ControlTipText ="Gives a description of the selected program"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2100
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =3390
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    Left =1695
                    Top =3120
                    Width =6780
                    Height =480
                    TabIndex =2
                    Name ="ObjectsUsed"
                    ControlSource ="strObjectsUsed"
                    FontName ="Tahoma"
                    ControlTipText ="A list of the objects used in building the sample."

                    LayoutCachedLeft =1695
                    LayoutCachedTop =3120
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =223
                            Left =300
                            Top =3120
                            Width =1380
                            Height =435
                            ForeColor =13209
                            Name ="Label14"
                            Caption ="Objects used:"
                            LayoutCachedLeft =300
                            LayoutCachedTop =3120
                            LayoutCachedWidth =1680
                            LayoutCachedHeight =3555
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =510
                    Top =630
                    Width =1125
                    Height =240
                    ForeColor =13209
                    Name ="Topic_Label"
                    Caption ="Program:"
                End
                Begin Label
                    OverlapFlags =215
                    Left =600
                    Top =2100
                    Width =1080
                    Height =240
                    ForeColor =13209
                    Name ="Label10"
                    Caption ="Description:"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ListWidth =1440
                    Left =2880
                    Top =150
                    Width =2760
                    Height =270
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboCategories"
                    RowSourceType ="Table/Query"
                    RowSource ="qrySelectTopic"
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Select a Category"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =150
                            Width =1620
                            Height =240
                            ForeColor =13209
                            Name ="ctrlCategories_Label"
                            Caption ="Select a Category:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    Left =960
                    Top =3705
                    Height =270
                    TabIndex =4
                    ForeColor =255
                    Name ="QNumber"
                    ControlSource ="pkeyQNumber"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =3705
                            Width =900
                            Height =240
                            Name ="Label7"
                            Caption ="QNumber:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =12311007
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =71
                    Left =2520
                    Top =60
                    Width =4155
                    Height =405
                    ForeColor =255
                    Name ="cmdShowExample"
                    Caption ="&Go"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    UnicodeAccessKey =71

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


Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.name
End Sub


Private Sub cmdShowExample_Click()
On Error GoTo cmdShowExample_Error

    If Not IsNull(Me![strObjectName]) Then

        Select Case Me![intObjectType]
        
            Case 1  'Run Reports
                DoCmd.OpenReport Me![strObjectName], acPreview
                
            Case 2  'Run Forms
                If Me![strObjectName] = "frmDepliant" Then
                    DoCmd.OpenForm Me![strObjectName], acFormDS
                ElseIf Me![strObjectName] = "fdlgSchedaFuelLogSearch" Then
                    DoCmd.OpenForm Me![strObjectName]
                    Forms!fdlgSchedaFuelLogSearch!cmdSearch.Visible = False
                    Forms!fdlgSchedaFuelLogSearch!CmdPrint.Visible = True
                Else
                    DoCmd.OpenForm Me![strObjectName]
                End If
            Case 3  'Run Macro
                ' in case you ever wanted to run macros.
                DoCmd.RunMacro Me![strObjectName]
                
            Case 4  'Run User Defined Function
                Eval (Me![strObjectName])
                
                
        End Select
    Else
        MsgBox "No Name in table"
    End If
cmdShowExample_Exit:
    On Error GoTo 0
    DoCmd.Close acForm, Me.name
    Exit Sub

cmdShowExample_Error:
    ErrorLog "cmdShowExample", Err, Error
    Resume cmdShowExample_Exit
End Sub

Private Sub cboCategories_AfterUpdate()
    Dim strSQL As String

    strSQL = "SELECT DISTINCTROW tblTopics.pkeyQNumber, tblTopics.strTopic, " & _
            "tblTopics.strSampleCategory AS Category FROM tblTopics " & _
            "WHERE (((tblTopics.strSampleCategory) = [Forms]![frmMenuReport]![cboCategories])) " & _
            "ORDER BY tblTopics.strTopic, tblTopics.strSampleCategory;"

    TopicList.RowSource = strSQL
    TopicList.Requery
    TopicList = TopicList.ItemData(1)
    TopicList_AfterUpdate
End Sub

Private Sub Form_Load()
   'Me.cboCategories.Value = "All Categories"
    Me.cboCategories.Value = "Report"
    TopicList = Me.TopicList.ItemData(0)
    cboCategories_AfterUpdate
End Sub

Private Sub TopicList_AfterUpdate()
    Dim rst As Object
    Set rst = Me.Recordset.Clone
    rst.FindFirst "[pkeyQNumber] = '" & Me![TopicList] & "'"
    Me.Bookmark = rst.Bookmark
End Sub
Sub btnClose_Click()
On Error GoTo Err_btnClose_Click
    DoCmd.Close
Exit_btnClose_Click:
    Exit Sub

Err_btnClose_Click:
    MsgBox Err.Description
    Resume Exit_btnClose_Click
    
End Sub

Private Sub TopicList_DblClick(Cancel As Integer)
    cmdShowExample_Click
End Sub
