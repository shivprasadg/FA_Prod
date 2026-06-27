Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11760
    DatasheetFontHeight =11
    ItemSuffix =744
    Left =5145
    Top =7245
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x26945c11f480e640
    End
    RecordSource ="SELECT TOP 100 ElectronicTitleRequest.tAssetID, ElectronicTitleRequest.SchGrp, E"
        "lectronicTitleRequest.MYear, ElectronicTitleRequest.Make, ElectronicTitleRequest"
        ".Lastsent, ElectronicTitleRequest.VIN, [Reports]![ElectronicTitleRequestOK]![Sig"
        "ned] AS Signed FROM ElectronicTitleRequest WHERE (((ElectronicTitleRequest.Sentd"
        "ate) Is Null)) ORDER BY ElectronicTitleRequest.Lastsent, ElectronicTitleRequest."
        "VIN;"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="Lastsent"
        End
        Begin BreakLevel
            ControlSource ="VIN"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =367
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =660
                    Top =30
                    Width =720
                    Height =315
                    Name ="Label691"
                    Caption ="MYear"
                    GroupTable =2
                    LayoutCachedLeft =660
                    LayoutCachedTop =30
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =345
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1440
                    Top =30
                    Width =2235
                    Height =315
                    Name ="Label699"
                    Caption ="Make"
                    GroupTable =2
                    LayoutCachedLeft =1440
                    LayoutCachedTop =30
                    LayoutCachedWidth =3675
                    LayoutCachedHeight =345
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =3735
                    Top =30
                    Width =2670
                    Height =315
                    Name ="Label706"
                    Caption ="VIN"
                    GroupTable =2
                    LayoutCachedLeft =3735
                    LayoutCachedTop =30
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =345
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =10425
                    Top =30
                    Width =1155
                    Height =315
                    Name ="Label704"
                    Caption ="FA AssetID"
                    GroupTable =2
                    LayoutCachedLeft =10425
                    LayoutCachedTop =30
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =345
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8925
                    Top =30
                    Width =1440
                    Height =315
                    Name ="Label701"
                    Caption ="Date Signed"
                    GroupTable =2
                    LayoutCachedLeft =8925
                    LayoutCachedTop =30
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =345
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =90
                    Top =30
                    Width =504
                    Height =315
                    Name ="Label733"
                    Caption ="Rec"
                    GroupTable =2
                    LayoutCachedLeft =90
                    LayoutCachedTop =30
                    LayoutCachedWidth =594
                    LayoutCachedHeight =345
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =6465
                    Top =30
                    Width =2399
                    Height =315
                    Name ="Label741"
                    Caption ="Sch | Grp"
                    GroupTable =2
                    LayoutCachedLeft =6465
                    LayoutCachedTop =30
                    LayoutCachedWidth =8864
                    LayoutCachedHeight =345
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =340
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =30
                    Width =504
                    Height =288
                    ForeColor =0
                    Name ="Text732"
                    ControlSource ="=1"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =564
                    LayoutCachedHeight =318
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =630
                    Top =30
                    Width =720
                    Height =288
                    TabIndex =1
                    ForeColor =0
                    Name ="Text672"
                    ControlSource ="MYear"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =630
                    LayoutCachedTop =30
                    LayoutCachedWidth =1350
                    LayoutCachedHeight =318
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1410
                    Top =30
                    Width =2235
                    Height =288
                    TabIndex =2
                    ForeColor =0
                    Name ="Text673"
                    ControlSource ="MAKE"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =1410
                    LayoutCachedTop =30
                    LayoutCachedWidth =3645
                    LayoutCachedHeight =318
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3705
                    Top =30
                    Width =2670
                    Height =288
                    TabIndex =3
                    ForeColor =0
                    Name ="Text671"
                    ControlSource ="VIN"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =3705
                    LayoutCachedTop =30
                    LayoutCachedWidth =6375
                    LayoutCachedHeight =318
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6435
                    Top =30
                    Width =2400
                    Height =288
                    TabIndex =4
                    ForeColor =0
                    Name ="Text738"
                    ControlSource ="SchGrp"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =6435
                    LayoutCachedTop =30
                    LayoutCachedWidth =8835
                    LayoutCachedHeight =318
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8895
                    Top =30
                    Height =288
                    TabIndex =5
                    ForeColor =0
                    Name ="Text675"
                    ControlSource ="Signed"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =8895
                    LayoutCachedTop =30
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =318
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =50
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10395
                    Top =30
                    Width =1152
                    Height =288
                    TabIndex =6
                    ForeColor =0
                    Name ="Text674"
                    ControlSource ="tAssetID"
                    FontName ="OCR A Extended"
                    GroupTable =3

                    LayoutCachedLeft =10395
                    LayoutCachedTop =30
                    LayoutCachedWidth =11547
                    LayoutCachedHeight =318
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =60
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin PageBreak
                    Top =60
                    Name ="PageBreak731"
                End
            End
        End
    End
End
CodeBehindForm
' See "ElectronicTitleRequestOK_HDR.cls"
