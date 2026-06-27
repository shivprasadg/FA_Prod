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
    Width =12870
    DatasheetFontHeight =11
    ItemSuffix =22
    Left =4845
    Top =9600
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x3017a3795687e640
    End
    RecordSource ="SELECT ChildParts.UnitREF, ChildParts.ChildVIN, [childyear] & \" \" & [childmake"
        "] & \" \" & [ChildClass] & \" \" & [childtype] AS ChildDesc, ChildParts.ChildInv"
        "oiceNum, ChildParts.childFinalCost, ChildParts.ChildVendorPytDate FROM ChildPart"
        "s ORDER BY ChildParts.ChildVIN DESC , [childyear] & \" \" & [childmake] & \" \" "
        "& [ChildClass] & \" \" & [childtype];"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =292
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =615
                    Width =60
                    Height =60
                    FontSize =9
                    Name ="UnitREF"
                    ControlSource ="UnitREF"
                    StatusBarText ="link to Unit"

                    LayoutCachedLeft =615
                    LayoutCachedWidth =675
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3150
                    Top =30
                    Width =3555
                    FontSize =9
                    TabIndex =2
                    ForeColor =7500402
                    Name ="Text1"
                    ControlSource ="ChildDesc"
                    GroupTable =2

                    LayoutCachedLeft =3150
                    LayoutCachedTop =30
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =270
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8475
                    Top =30
                    Width =1110
                    FontSize =9
                    TabIndex =4
                    ForeColor =7500402
                    Name ="childFinalCost"
                    ControlSource ="childFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =2

                    LayoutCachedLeft =8475
                    LayoutCachedTop =30
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1140
                    Top =30
                    Width =1950
                    FontSize =9
                    TabIndex =1
                    ForeColor =7500402
                    Name ="ChildVIN"
                    ControlSource ="ChildVIN"
                    StatusBarText ="obsoleted 2-10-10"
                    GroupTable =2

                    LayoutCachedLeft =1140
                    LayoutCachedTop =30
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =270
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6765
                    Top =30
                    Width =1650
                    FontSize =9
                    TabIndex =3
                    ForeColor =7500402
                    Name ="ChildInvoiceNum"
                    ControlSource ="ChildInvoiceNum"
                    GroupTable =2

                    LayoutCachedLeft =6765
                    LayoutCachedTop =30
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9645
                    Top =30
                    Width =840
                    FontSize =9
                    TabIndex =5
                    ForeColor =7500402
                    Name ="ChildVendorPytDate"
                    ControlSource ="ChildVendorPytDate"
                    Format ="mm/dd/yy"
                    GroupTable =2
                    ShowDatePicker =1

                    LayoutCachedLeft =9645
                    LayoutCachedTop =30
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =270
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
            End
        End
    End
End
