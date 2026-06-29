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
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10566
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =4635
    Top =2250
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x9b88e0bc8b39e540
    End
    RecordSource ="SELECT vw_SixKeys.ClientShortName, vw_SixKeys.MLNo, vw_SixKeys.AssetID, vw_SixKe"
        "ys.VIN, UnitTitleLog.UnitTitleLogDate, UnitTitleLog.TitleNumber, EntryType.Entry"
        "TypeDescription, UnitTitleLog.EntryReason AS Reason, UnitTitleLog.TitleSt AS Tit"
        "leState, UnitTitleLog.ExpectedReturn, IIf([EntryType].[entrytypeid]=1,Null,IIf(["
        "ExpectedReturn]<>0,\"Yes\",\"Final\")) AS RTN, UnitTitleLog.UnitTitleLogID FROM "
        "vw_SixKeys LEFT JOIN ((UnitTitleLog LEFT JOIN Reasons ON UnitTitleLog.EntryReaso"
        "nId = Reasons.ReasonId) LEFT JOIN EntryType ON UnitTitleLog.EntryTypeId = EntryT"
        "ype.EntryTypeId) ON vw_SixKeys.AssetID = UnitTitleLog.UnitRef WHERE (((vw_SixKey"
        "s.AssetID)=[forms]![unitsform]![unitid])) ORDER BY UnitTitleLog.UnitTitleLogDate"
        " DESC , UnitTitleLog.UnitTitleLogID DESC;"
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
        Begin Rectangle
            BorderLineStyle =0
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
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="UnitTitleLogID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1035
            Name ="ReportHeader"
            Begin
                Begin Rectangle
                    Left =15
                    Top =480
                    Width =9750
                    Height =135
                    BackColor =-2147483617
                    Name ="Box10"
                    LayoutCachedLeft =15
                    LayoutCachedTop =480
                    LayoutCachedWidth =9765
                    LayoutCachedHeight =615
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =630
                    Width =9720
                    Height =405
                    FontSize =14
                    Name ="Text4"
                    ControlSource ="=\"Log for  Asset ID: \" & [AssetID] & \"| VIN: \" & [VIN]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =630
                    LayoutCachedWidth =9750
                    LayoutCachedHeight =1035
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    Left =5955
                    Top =60
                    Width =3780
                    Height =585
                    FontSize =24
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label9"
                    Caption ="Unit Title Log"
                    LayoutCachedLeft =5955
                    LayoutCachedTop =60
                    LayoutCachedWidth =9735
                    LayoutCachedHeight =645
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    Left =6540
                    Top =45
                    Width =3210
                    Height =585
                    FontSize =24
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Unit Title Log"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =45
                    LayoutCachedWidth =9750
                    LayoutCachedHeight =630
                End
            End
        End
        Begin PageHeader
            Height =362
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =36
                    Top =36
                    Width =1296
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =-2147483617
                    Name ="Label0"
                    Caption ="Entry Date:"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =324
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =1404
                    Top =36
                    Width =876
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483617
                    ForeColor =-2147483617
                    Name ="Label1"
                    Caption ="In/Out:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =1404
                    LayoutCachedTop =36
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =324
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =4944
                    Top =36
                    Width =4080
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =-2147483617
                    Name ="Label2"
                    Caption ="Reason:"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =4944
                    LayoutCachedTop =36
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =324
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =2352
                    Top =36
                    Width =1752
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =-2147483617
                    Name ="Label3"
                    Caption ="TitleNumber:"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =2352
                    LayoutCachedTop =36
                    LayoutCachedWidth =4104
                    LayoutCachedHeight =324
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =4164
                    Top =36
                    Width =720
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =-2147483617
                    Name ="Label13"
                    Caption ="State"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =4164
                    LayoutCachedTop =36
                    LayoutCachedWidth =4884
                    LayoutCachedHeight =324
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =9084
                    Top =36
                    Width =1452
                    Height =288
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =-2147483617
                    Name ="Label14"
                    Caption ="Return"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =9084
                    LayoutCachedTop =36
                    LayoutCachedWidth =10536
                    LayoutCachedHeight =324
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =362
            Name ="Detail"
            AlternateBackColor =15527148
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =1296
                    Height =288
                    FontSize =10
                    ForeColor =-2147483617
                    Name ="UnitTitleLogDate"
                    ControlSource ="UnitTitleLogDate"
                    Format ="General Date"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1404
                    Top =36
                    Width =876
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =-2147483617
                    ForeColor =-2147483617
                    Name ="EntryTypeDescription"
                    ControlSource ="EntryTypeDescription"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1404
                    LayoutCachedTop =36
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4944
                    Top =36
                    Width =4080
                    Height =288
                    ColumnWidth =3465
                    FontSize =10
                    TabIndex =4
                    ForeColor =-2147483617
                    Name ="ReasonDescription"
                    ControlSource ="Reason"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4944
                    LayoutCachedTop =36
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2352
                    Top =36
                    Width =1752
                    Height =288
                    ColumnWidth =1605
                    FontSize =10
                    TabIndex =2
                    ForeColor =-2147483617
                    Name ="TitleNumber"
                    ControlSource ="TitleNumber"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2352
                    LayoutCachedTop =36
                    LayoutCachedWidth =4104
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4164
                    Top =36
                    Width =720
                    Height =288
                    FontSize =10
                    TabIndex =3
                    ForeColor =-2147483617
                    Name ="Text12"
                    ControlSource ="TitleState"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4164
                    LayoutCachedTop =36
                    LayoutCachedWidth =4884
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9084
                    Top =36
                    Width =1452
                    Height =288
                    FontSize =10
                    TabIndex =5
                    ForeColor =-2147483617
                    Name ="Text15"
                    ControlSource ="RTN"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =9084
                    LayoutCachedTop =36
                    LayoutCachedWidth =10536
                    LayoutCachedHeight =324
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin PageFooter
            Height =195
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2730
                    Height =195
                    FontSize =8
                    FontWeight =700
                    Name ="Text6"
                    ControlSource ="=\"printed on \" & Date() & \" @ \" & Time()"

                    LayoutCachedWidth =2730
                    LayoutCachedHeight =195
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7774
                    Width =1755
                    Height =195
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="Text8"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =7774
                    LayoutCachedWidth =9529
                    LayoutCachedHeight =195
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
