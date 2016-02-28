Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4800
    DatasheetFontHeight =10
    Left =3585
    Top =135
    Right =8355
    Bottom =2865
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3e6c23138a38e240
    End
    OnOpen ="[Event Procedure]"
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
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
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
            OldBorderStyle =0
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
        Begin Section
            CanGrow = NotDefault
            Height =2760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =90
                    Top =90
                    Width =4620
                    Height =660
                    BackColor =8421504
                    Name ="Box0"
                End
                Begin Label
                    OverlapFlags =223
                    Left =270
                    Top =225
                    Width =3360
                    Height =480
                    FontSize =18
                    ForeColor =3355443
                    Name ="Label1"
                    Caption ="Current Date/Time"
                End
                Begin Label
                    OverlapFlags =223
                    Left =240
                    Top =195
                    Width =3360
                    Height =480
                    FontSize =18
                    ForeColor =16777215
                    Name ="Label2"
                    Caption ="Current Date/Time"
                End
                Begin Rectangle
                    SpecialEffect =1
                    OverlapFlags =223
                    Left =30
                    Top =30
                    Width =4740
                    Height =2700
                    Name ="Box3"
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =90
                    Top =810
                    Width =4620
                    Height =1860
                    Name ="Box1"
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =270
                    Top =990
                    Width =4320
                    Height =1080
                    Name ="Box2"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =2190
                    Top =1110
                    Width =2280
                    Height =360
                    FontSize =10
                    Name ="txtDate"
                    FontName ="Arial"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =330
                            Top =1110
                            Width =1800
                            Height =360
                            FontSize =12
                            FontWeight =800
                            Name ="lblSearch"
                            Caption ="Date:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =2190
                    Top =1590
                    Width =2280
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="txtTime"
                    FontName ="Arial"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =330
                            Top =1590
                            Width =1800
                            Height =360
                            FontSize =12
                            FontWeight =800
                            Name ="Label23"
                            Caption ="Time:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =75
                    Left =1365
                    Top =2145
                    Width =840
                    Height =420
                    TabIndex =2
                    Name ="CmdOK"
                    Caption ="O&K"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =2445
                    Top =2145
                    Width =840
                    Height =420
                    TabIndex =3
                    Name ="CmdApply"
                    Caption ="Apply"
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
Option Explicit

Private Sub CmdApply_Click()
On Error GoTo err_CmdApply_Click

    If IsNull(Me!txtDate) Or IsNull(Me!txtTime) Then
        MsgBox IIf(IsNull(Me!txtDate), "Date is empty.", _
            "Time is empty."), vbInformation
        Exit Sub
    End If
    Date = txtDate
    Time = txtTime
    Me.CmdOK.SetFocus
    Me.CmdApply.Enabled = False
    
exit_err_CmdApply_Click:
    Exit Sub
    
err_CmdApply_Click:
    MsgBox Err.Description, vbInformation
    Resume exit_err_CmdApply_Click:
End Sub

Private Sub CmdOK_Click()
On Error GoTo err_CmdOK
    DoCmd.Close
    
exit_err_CmdOK:
    Exit Sub
    
err_CmdOK:
    MsgBox Err.Description, vbInformation
    Resume exit_err_CmdOK
End Sub

Private Sub Form_Open(Cancel As Integer)
    txtDate = Date
    txtTime = Time()
End Sub

Private Sub txtDate_Change()
    Me!CmdApply.Enabled = True
End Sub

Private Sub txtTime_Change()
    Me!CmdApply.Enabled = True
End Sub
