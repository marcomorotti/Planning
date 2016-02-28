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
    Top =4170
    Right =15075
    Bottom =6180
    HelpContextId =52
    RecSrcDt = Begin
        0xa3aa30f84e40e140
    End
    Caption ="Articoli Obsoleti"
    OnOpen ="[Event Procedure]"
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
                    Caption ="Articoli OBSOLETI"
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
                    Name ="cmdExcel"
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
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3120
                    Width =381
                    Height =366
                    FontSize =9
                    TabIndex =3
                    HelpContextId =52
                    ForeColor =10040115
                    Name ="cmdHelp"
                    Caption ="Help"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
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

                    LayoutCachedLeft =3120
                    LayoutCachedWidth =3501
                    LayoutCachedHeight =366
                End
            End
        End
        Begin Section
            Height =1455
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1719
                    Top =283
                    Width =576
                    Height =300
                    FontSize =11
                    ForeColor =255
                    Name ="Wks"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Arial"

                    LayoutCachedLeft =1719
                    LayoutCachedTop =283
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =583
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =135
                            Width =1665
                            Height =600
                            ForeColor =255
                            Name ="Text14"
                            Caption ="1-INSERISCI Mesi\015\012Obsolescenza:"
                            LayoutCachedTop =135
                            LayoutCachedWidth =1665
                            LayoutCachedHeight =735
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    Left =3720
                    Top =135
                    Height =300
                    FontSize =11
                    TabIndex =1
                    Name ="FromDate"
                    Format ="Short Date"
                    FontName ="Arial"
                    ShowDatePicker =0

                    LayoutCachedLeft =3720
                    LayoutCachedTop =135
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =435
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2625
                            Top =135
                            Width =1095
                            Height =240
                            Name ="Text2"
                            Caption ="Data Inizio:"
                            LayoutCachedLeft =2625
                            LayoutCachedTop =135
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    Left =3720
                    Top =855
                    Height =300
                    FontSize =11
                    TabIndex =2
                    Name ="ToDate"
                    Format ="Short Date"
                    FontName ="Arial"
                    ShowDatePicker =0

                    LayoutCachedLeft =3720
                    LayoutCachedTop =855
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1155
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2325
                            Top =855
                            Width =1395
                            Height =240
                            Name ="Text4"
                            Caption ="Data Fine:"
                            LayoutCachedLeft =2325
                            LayoutCachedTop =855
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =855
                    Width =1755
                    Height =600
                    ForeColor =255
                    Name ="Etichetta22"
                    Caption ="2-Premi tasto Excel per esportare i dati"
                    LayoutCachedLeft =60
                    LayoutCachedTop =855
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1455
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
If IsNothing(Me!FromDate) Or IsNothing(Me!ToDate) Then
    MsgBox "Devi inserire i Mesi di Obsolescenza richiesti!!!", vbInformation, _
        gstrAppTitle
    Me.Wks.SetFocus
    Exit Sub
End If
Dim Response As Integer
  
    

' Dim nRecords As Integer 'Long
On Error GoTo Err_cmdExcel_Click:

Response = MsgBox("Vuoi calcolare gli ARTICOLI OBSOLETI ? ", vbYesNo, _
    "Continue")
If Response = vbYes Then

    Dim AnnoI As String
    Dim MeseI As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim AnnoOggi As String
    Dim MeseOggi As String
    Dim strDataLine As String
    Dim intFile As Integer
    Dim filenm As String
    Dim i As Integer
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim db0 As Database
    Dim conn As ADODB.Connection
    Dim Articolo As Variant

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    
    ' Memorizzo le date
    AnnoI = Year(Me.FromDate)
    MeseI = Month(Me.FromDate)
    AnnoF = Year(Me.ToDate)
    MeseF = Month(Me.ToDate)
    AnnoOggi = Year(Date)
    MeseOggi = Month(Date)
    ' Crea File di output
    intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "ArticoliObsoleti_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
    & "ArticoliObsoleti_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
            & "ArticoliObsoleti_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "ArticoliObsoleti_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrAppTitle) Then
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Codice" & Chr(9) & "Descrizione" & Chr(9) & "Giacenza" & Chr(9) & "Costo CSC"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
        
    
    Set bqry = db0.QueryDefs("qryArticoliObsoleti")
    
    ' Imposta i parametri qry parametrica
'    bqry.Parameters("AnnoI").Value = AnnoI
'    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
'    bqry.Parameters("AnnoOggi").Value = AnnoOggi
'    bqry.Parameters("MeseOggi").Value = MeseOggi
    Set brs = bqry.OpenRecordset
    Articolo = ""
    Do While Not brs.EOF
    If brs.Fields("Cod_art").Value <> Articolo Then
        strDataLine = brs.Fields("Cod_art").Value & Chr(9) & _
            brs.Fields("Des_art").Value & Chr(9) & brs.Fields("Giac_Media").Value & Chr(9) & _
            brs.Fields("Cs_Csc").Value
        Print #intFile, strDataLine
    End If
    Articolo = brs.Fields("Cod_art").Value
    brs.MoveNext
    strDataLine = ""
    Loop
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "ArticoliObsoleti_" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set db0 = Nothing
    Close #intFile
Else
Exit Sub
End If
End If
Exit_cmdExcel_Click:
    Exit Sub

Err_cmdExcel_Click:
    MsgBox Err.Description
    Resume Exit_cmdExcel_Click
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdHelpReportCost_Click()
DoCmd.OpenForm "z Help Text for User", , , "zhID = 90"
End Sub


Private Sub cmdHelp_Click()
Dim FormHelpId As Long
Dim curForm As Form
'Set the curForm variable to the currently active form.
Set curForm = Screen.ActiveForm
FormHelpId = 52
'Call the function to start the Help file, passing it the name of the
'Help file and context-id.
Show_Help FormHelpFile, FormHelpId
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim db0 As Database ' ******* DATA BASE CORRENTE ***********
Dim rs As Recordset
Dim conn As ADODB.Connection
Dim intI As Double
' Verificare i Dim
Dim intMonth As Integer, intYear As Integer, intStartDate As Variant, _
    intStartDateS As Variant, intEndDate As Variant, intLastDay As Integer, NumMesiC As Integer
Dim DatiGen As New ADODB.Recordset
    Dim AnnoC As String
    Dim MeseC As String
    Dim AnnoI As String
    Dim MeseI As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim Cmd As New ADODB.Command
Set db0 = CurrentDb
Set conn = CurrentProject.Connection
'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
'    NumMesiC = DatiGen.Fields("UpdateMesi_consumo")
'    intYear = DatiGen.Fields("UpdateAnno_Calcolo")
'    intMonth = DatiGen.Fields("UpdateMese_Calcolo")
'    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
'    intStartDate = DateSerial(intYear, intMonth, 1)
'    intStartDate = Format(intStartDate, "dd/mmm/yyyy")
'    AnnoC = Year(intStartDate)
intYear = DatiGen.Fields("Anno_calcolo")
intMonth = DatiGen.Fields("Mese_calcolo")
intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
intStartDate = DateSerial(intYear, intMonth, intLastDay)
intStartDate = Format(intStartDate, "dd/mmm/yyyy")
Me.FromDate = intStartDate
End Sub

Private Sub Wks_AfterUpdate()
    If (IsNothing(Me!FromDate) Or Me!Wks = 0) Then
        MsgBox "Devi inserire un numero > 0", vbInformation, gstrAppTitle
        Exit Sub
    End If
    Dim intLastDay As Integer, intEndDate As Variant, NumMesiC As Integer, intStartDate As Variant
    Dim intYear As Integer, intMonth As Integer
    Dim db0 As Database
    Dim DatiGen As New ADODB.Recordset
    Dim conn As ADODB.Connection
    
    intYear = Year(Me.FromDate)
    intMonth = Month(Me.FromDate)
    intEndDate = DateAdd("m", -Me!Wks + 2, DateSerial(intYear, intMonth, 1))
    
' ********* Verifico che non si vada fuori scala
Set db0 = CurrentDb
Set conn = CurrentProject.Connection
'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    NumMesiC = DatiGen.Fields("UpdateMesi_consumo")
    intYear = DatiGen.Fields("UpdateAnno_Calcolo")
    intMonth = DatiGen.Fields("UpdateMese_Calcolo")
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    intStartDate = DateSerial(intYear, intMonth, 1)
      If DateDiff("d", intEndDate, intStartDate) > 0 Then
        MsgBox "Devi inserire un numero <= " & NumMesiC, vbInformation, gstrAppTitle
        Exit Sub
    End If
  
' AGGIORNA CAMPO DELLA FORM
Me!ToDate = Format(intEndDate, "dd/mmm/yyyy")
End Sub
