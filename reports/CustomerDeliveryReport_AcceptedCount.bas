Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    PicturePages =1
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6930
    DatasheetFontHeight =11
    ItemSuffix =5
    Left =6300
    Top =10215
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xab230c29cfcbe440
    End
    RecordSource ="SELECT [DelAccptCntSummary].[Accpt], [DelAccptCntSummary].[UnAccpt] FROM [DelAcc"
        "ptCntSummary]"
    Caption ="DelAccptCntSummary subreport"
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
            TextFontCharSet =238
            FontSize =10
            BorderColor =-2147483609
            ForeColor =-2147483615
            FontName ="Calibri"
            GridlineColor =-2147483613
        End
        Begin Rectangle
            BorderLineStyle =0
            BackColor =-2147483616
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin Line
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin Image
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackColor =-2147483616
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin CommandButton
            TextFontCharSet =238
            FontWeight =400
            FontName ="Calibri"
            GridlineColor =-2147483613
            BorderLineStyle =0
        End
        Begin OptionButton
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            GridlineColor =-2147483613
        End
        Begin OptionGroup
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =238
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            FontSize =10
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
            GridlineColor =-2147483613
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontCharSet =238
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =10
            BorderColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483613
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =238
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            FontSize =10
            BorderColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483613
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin UnboundObjectFrame
            BackStyle =0
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
            GridlineColor =-2147483613
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontWeight =400
            FontName ="Calibri"
            GridlineColor =-2147483613
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            BackStyle =0
            FontSize =10
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            GridlineColor =-2147483613
            LabelX =-1800
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =375
            Name ="Detail"
            AlternateBackColor =-2147483613
            Begin
                Begin TextBox
                    TextFontCharSet =204
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2670
                    Width =1155
                    Height =285
                    FontSize =11
                    FontWeight =700
                    Name ="UnAccpt"
                    ControlSource ="UnAccpt"

                    LayoutCachedLeft =2670
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =285
                    Begin
                        Begin Label
                            TextFontCharSet =204
                            Left =825
                            Width =1800
                            Height =285
                            FontSize =11
                            FontWeight =700
                            ForeColor =-2147483617
                            Name ="Label4"
                            Caption ="Total Unaccepted:"
                            LayoutCachedLeft =825
                            LayoutCachedWidth =2625
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AutoHeight =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
        End
    End
End
