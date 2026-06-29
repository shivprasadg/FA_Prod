Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5970
    DatasheetFontHeight =11
    ItemSuffix =434
    Left =7635
    Top =6615
    Right =13620
    Bottom =8700
    RecSrcDt = Begin
        0x33536e62d5fee440
    End
    RecordSource ="SELECT * FROM Schgrp"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
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
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =2085
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =360
                    Top =1605
                    Width =798
                    Height =276
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =12
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    TopPadding =0

                    LayoutCachedLeft =360
                    LayoutCachedTop =1605
                    LayoutCachedWidth =1158
                    LayoutCachedHeight =1881
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =36
                    Top =36
                    Width =2832
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    Name ="Label474"
                    Caption ="Vendor"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =276
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2940
                    Top =36
                    Width =1200
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    Name ="Label475"
                    Caption ="PO Sent Date"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =2940
                    LayoutCachedTop =36
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =276
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4212
                    Top =36
                    Width =1680
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    Name ="Label480"
                    Caption ="PO #"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =4212
                    LayoutCachedTop =36
                    LayoutCachedWidth =5892
                    LayoutCachedHeight =276
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =24
                    ListWidth =5904
                    Left =36
                    Top =348
                    Width =2832
                    FontSize =9
                    BorderColor =12632256
                    ForeColor =0
                    Name ="POVendor"
                    ControlSource ="POVendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                        " ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;1872;2160"
                    StatusBarText ="link to Vendor Table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =36
                    LayoutCachedTop =348
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =588
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =348
                    Width =1200
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =12632256
                    ForeColor =0
                    Name ="POSentDt"
                    ControlSource ="POSentDt"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =348
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =588
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4212
                    Top =348
                    Width =1680
                    FontSize =9
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PONum1"
                    ControlSource ="PONum1"
                    Format ="mm/dd/yy"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =4212
                    LayoutCachedTop =348
                    LayoutCachedWidth =5892
                    LayoutCachedHeight =588
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =24
                    ListWidth =5904
                    Left =36
                    Top =660
                    Width =2832
                    FontSize =9
                    TabIndex =3
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO2Vendor"
                    ControlSource ="PO2Vendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                        " ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;1872;2160"
                    StatusBarText ="link to Vendor Table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =36
                    LayoutCachedTop =660
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =660
                    Width =1200
                    FontSize =9
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO2SentDt"
                    ControlSource ="PO2SentDt"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =660
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4212
                    Top =660
                    Width =1680
                    FontSize =9
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PONum2"
                    ControlSource ="PONum2"
                    Format ="mm/dd/yy"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =4212
                    LayoutCachedTop =660
                    LayoutCachedWidth =5892
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =24
                    ListWidth =5904
                    Left =36
                    Top =972
                    Width =2832
                    FontSize =9
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO3Vendor"
                    ControlSource ="PO3Vendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                        " ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;1872;2160"
                    StatusBarText ="link to Vendor Table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =36
                    LayoutCachedTop =972
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =1212
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =972
                    Width =1200
                    FontSize =9
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO3SentDt"
                    ControlSource ="PO3SentDt"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =972
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1212
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4212
                    Top =972
                    Width =1680
                    FontSize =9
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PONum3"
                    ControlSource ="PONum3"
                    Format ="mm/dd/yy"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =4212
                    LayoutCachedTop =972
                    LayoutCachedWidth =5892
                    LayoutCachedHeight =1212
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =24
                    ListWidth =5904
                    Left =36
                    Top =1284
                    Width =2832
                    FontSize =9
                    TabIndex =9
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO4Vendor"
                    ControlSource ="PO4Vendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                        " ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;1872;2160"
                    StatusBarText ="link to Vendor Table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =36
                    LayoutCachedTop =1284
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =1524
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =1284
                    Width =1200
                    FontSize =9
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PO4SentDt"
                    ControlSource ="PO4SentDt"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1284
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1524
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4212
                    Top =1284
                    Width =1680
                    FontSize =9
                    TabIndex =11
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =14277081
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PONum4"
                    ControlSource ="PONum4"
                    Format ="mm/dd/yy"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="StayOpen"
                    GroupTable =2

                    LayoutCachedLeft =4212
                    LayoutCachedTop =1284
                    LayoutCachedWidth =5892
                    LayoutCachedHeight =1524
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "GroupPOSub.cls"
