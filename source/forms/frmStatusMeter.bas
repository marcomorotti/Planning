Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =1
    Width =3168
    ItemSuffix =7
    Left =10965
    Top =4560
    Right =14130
    Bottom =5625
    RecSrcDt = Begin
        0x05e0b3edf211e240
    End
    Caption ="Status Meter"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            SpecialEffect =2
            BackStyle =0
            TextAlign =2
            ForeColor =255
            FontName ="System"
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
            Width =0
            Height =360
            BackColor =65535
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
            Width =1152
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
            Height =1080
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =144
                    Top =120
                    Height =285
                    Name ="recStatus"
                End
                Begin Label
                    OverlapFlags =87
                    Left =144
                    Top =120
                    Width =2880
                    Height =285
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="0% Completed"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =1152
                    Top =600
                    Width =864
                    Name ="cmdCancel"
                    Caption ="&Cancel"
                    OnClick ="[Event Procedure]"

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
Public Property Let InitMeter(fIncludeCancel As Boolean, strTitle As String)

    Me!recStatus.Width = 0
    Me!lblStatus.Caption = "0% complete"
    Me.Caption = strTitle
    '   Me!cmdCancel.Visible = fIncludeCancel
    If fIncludeCancel Then
        Me!cmdCancel.Visible = True
        Me!cmdCancel.Enabled = True
    Else
        Me!cmdCancel.Enabled = False
        Me!cmdCancel.Visible = False
    End If

    DoCmd.RepaintObject

    mfCancel = False

End Property

Public Property Let UpdateMeter(intValue As Integer)

   Me!recStatus.Width = CInt(Me!lblStatus.Width * (intValue / 100))
   Me!lblStatus.Caption = Format$(intValue, "##") & "% complete"

   DoCmd.RepaintObject

End Property
 
Public Property Get Cancelled() As Boolean
   Cancelled = mfCancel
End Property

Private Sub cmdCancel_Click()
    mfCancel = True
    ' 20150423 Mrco
    DoCmd.Close acForm, Me.name
End Sub
