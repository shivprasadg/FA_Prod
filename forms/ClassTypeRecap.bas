Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6705
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =5190
    Top =6780
    Right =14145
    Bottom =12675
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x82b966d99610e440
    End
    RecordSource ="SELECT DISTINCT CLASS.TITLED, CLASS.CLASS, lookup_groupTypes.groupType, lookup_g"
        "roupTypes.GrpTypAbrev FROM CLASS INNER JOIN lookup_groupTypes ON CLASS.CLASS = l"
        "ookup_groupTypes.grpClass ORDER BY CLASS.TITLED, CLASS.CLASS, lookup_groupTypes."
        "groupType;"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
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
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin PageHeader
            DisplayWhen =1
            Height =345
            Name ="PageHeaderSection"
        End
        Begin Section
            Height =270
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =85
                    Left =180
                    Top =30
                    Name ="TITLED"
                    ControlSource ="TITLED"

                    LayoutCachedLeft =180
                    LayoutCachedTop =30
                    LayoutCachedWidth =440
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =470
                            Width =615
                            Height =240
                            FontSize =9
                            Name ="Label0"
                            Caption ="TITLED"
                            LayoutCachedLeft =470
                            LayoutCachedWidth =1085
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1380
                    Width =1380
                    ColumnWidth =1950
                    FontSize =9
                    TabIndex =1
                    Name ="CLASS"
                    ControlSource ="CLASS"
                    StatusBarText ="CLASS ASSOCIATED WITH UNIT TYPE FOR REPORTING"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Class is required"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2820
                    Top =15
                    Width =2280
                    ColumnWidth =3930
                    FontSize =9
                    TabIndex =2
                    Name ="groupType"
                    ControlSource ="groupType"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =15
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5145
                    Height =255
                    FontSize =9
                    TabIndex =3
                    Name ="GrpTypAbrev"
                    ControlSource ="GrpTypAbrev"

                    LayoutCachedLeft =5145
                    LayoutCachedWidth =6585
                    LayoutCachedHeight =255
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
        End
    End
End
