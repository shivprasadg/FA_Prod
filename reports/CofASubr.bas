Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8442
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =6780
    Top =7695
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xe47a9f2b1880e440
    End
    RecordSource ="CofAChildSub"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin BreakLevel
            ControlSource ="UnitREF"
        End
        Begin BreakLevel
            ControlSource ="ChildType"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =194
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8370
                    Width =72
                    Height =0
                    Name ="UnitREF"
                    ControlSource ="UnitREF"
                    StatusBarText ="link to Unit"

                    LayoutCachedLeft =8370
                    LayoutCachedWidth =8442
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =15
                    Width =8340
                    Height =194
                    FontSize =9
                    TabIndex =1
                    Name ="Text11"
                    ControlSource ="=IIf(IsNull([ChildDesc]),Null,[ChildDesc])"

                    LayoutCachedLeft =15
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =194
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "CofASubr.cls"
