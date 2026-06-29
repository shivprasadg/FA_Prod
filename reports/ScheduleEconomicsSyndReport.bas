Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15360
    DatasheetFontHeight =11
    ItemSuffix =2557
    Left =5670
    Top =3045
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x695dcb0355f5e540
    End
    RecordSource ="Select * FROM vw_EconoProfileGroup Where SchId=2228"
    OnOpen ="[Event Procedure]"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Client"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Schedule"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="GroupNum"
        End
        Begin PageHeader
            Height =1875
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9705
                    Width =5595
                    Height =225
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =700
                    Name ="Text1490"
                    ControlSource ="=Now()"

                    LayoutCachedLeft =9705
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9705
                    Top =225
                    Width =5595
                    Height =225
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="Text1515"
                    ControlSource ="=\" Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =9705
                    LayoutCachedTop =225
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =450
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =975
                    Width =8985
                    Height =330
                    ColumnWidth =3540
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    ForeColor =10040879
                    Name ="Client"
                    ControlSource ="ClientGroupWClientName"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =60
                    LayoutCachedTop =975
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =1305
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            Left =60
                            Top =795
                            Width =2775
                            Height =210
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label0"
                            Caption ="Client Group | Lessee"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =795
                            LayoutCachedWidth =2835
                            LayoutCachedHeight =1005
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =1515
                    Width =8985
                    Height =360
                    ColumnWidth =3315
                    ColumnOrder =3
                    TabIndex =5
                    ForeColor =0
                    Name ="Guarantor"
                    ControlSource ="Guarantor"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =60
                    LayoutCachedTop =1515
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =1875
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            Left =60
                            Top =1290
                            Width =885
                            Height =210
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label1"
                            Caption ="Guarantor"
                            FontName ="Segoe UI"
                            BottomPadding =150
                            LayoutCachedLeft =60
                            LayoutCachedTop =1290
                            LayoutCachedWidth =945
                            LayoutCachedHeight =1500
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9075
                    Top =975
                    Width =6105
                    Height =330
                    ColumnWidth =2805
                    ColumnOrder =1
                    TabIndex =3
                    ForeColor =0
                    Name ="Assignee"
                    ControlSource ="BankName"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =9075
                    LayoutCachedTop =975
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =1305
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =9075
                            Top =765
                            Width =1500
                            Height =270
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2"
                            Caption ="Assignee Bank"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =9075
                            LayoutCachedTop =765
                            LayoutCachedWidth =10575
                            LayoutCachedHeight =1035
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9075
                    Top =1515
                    Width =3795
                    Height =360
                    ColumnWidth =1725
                    ColumnOrder =2
                    TabIndex =4
                    ForeColor =0
                    Name ="MasterLease"
                    ControlSource ="mlno"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =9075
                    LayoutCachedTop =1515
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =1875
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =9075
                            Top =1290
                            Width =1635
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label3"
                            Caption ="Master Lease No."
                            FontName ="Segoe UI"
                            LayoutCachedLeft =9075
                            LayoutCachedTop =1290
                            LayoutCachedWidth =10710
                            LayoutCachedHeight =1530
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin Image
                    SizeMode =1
                    PictureType =2
                    Left =6990
                    Width =2034
                    Height =939
                    BorderColor =0
                    Name ="Auto_Logo0"
                    GridlineColor =0

                    LayoutCachedLeft =6990
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =939
                    TabIndex =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =105
                    Width =6930
                    Height =480
                    FontSize =16
                    FontWeight =700
                    ForeColor =0
                    Name ="ReportHeader"
                    Caption ="Final Syndication Economics"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =60
                    LayoutCachedTop =105
                    LayoutCachedWidth =6990
                    LayoutCachedHeight =585
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12870
                    Top =1515
                    Width =2310
                    Height =345
                    TabIndex =6
                    ForeColor =0
                    Name ="Text2219"
                    ControlSource ="FAID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12870
                    LayoutCachedTop =1515
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =1860
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =12870
                            Top =1290
                            Width =1635
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2220"
                            Caption ="FAID"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =12870
                            LayoutCachedTop =1290
                            LayoutCachedWidth =14505
                            LayoutCachedHeight =1530
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader2"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =983
            BreakLevel =2
            Name ="GroupHeader1"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Top =465
                    Width =15300
                    Height =518
                    BackColor =14277081
                    Name ="Box532"
                    LayoutCachedTop =465
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =983
                    BackShade =85.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =14970
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    ForeColor =10040879
                    Name ="GroupHeader"
                    ControlSource ="=\"Schedule \" & [Schedule] & \" - Group \" & [GroupNum] & \" - \" & [GroupDesc]"
                    FontName ="Segoe UI"
                    GroupTable =50
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =450
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =50
                End
                Begin Label
                    TextAlign =2
                    Left =672
                    Top =480
                    Width =1380
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label7"
                    Caption ="Cost \015\012(per unit)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =672
                    LayoutCachedTop =480
                    LayoutCachedWidth =2052
                    LayoutCachedHeight =975
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =3552
                    Top =480
                    Width =1308
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label8"
                    Caption ="Synd Sale Price\015\012(avg per unit)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3552
                    LayoutCachedTop =480
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =975
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =36
                    Top =480
                    Width =576
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label9"
                    Caption ="# Units"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =480
                    LayoutCachedWidth =612
                    LayoutCachedHeight =975
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =7128
                    Top =480
                    Width =780
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label12"
                    Caption ="MPY\015\012Allowed"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7128
                    LayoutCachedTop =480
                    LayoutCachedWidth =7908
                    LayoutCachedHeight =975
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =7968
                    Top =480
                    Width =576
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label13"
                    Caption ="Term mos."
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7968
                    LayoutCachedTop =480
                    LayoutCachedWidth =8544
                    LayoutCachedHeight =975
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =10620
                    Top =480
                    Width =1272
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label19"
                    Caption ="RV $\015\012(per unit)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10620
                    LayoutCachedTop =480
                    LayoutCachedWidth =11892
                    LayoutCachedHeight =975
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =11892
                    Top =480
                    Width =648
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label18"
                    Caption ="RV %"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11892
                    LayoutCachedTop =480
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =975
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13350
                    Top =45
                    Width =765
                    ColumnWidth =1455
                    FontSize =9
                    TabIndex =1
                    Name ="GroupID"
                    ControlSource ="GroupID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13350
                    LayoutCachedTop =45
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =13932
                    Top =480
                    Width =672
                    Height =495
                    FontSize =8
                    Name ="Label26"
                    Caption ="Synd\015\012 %"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13932
                    LayoutCachedTop =480
                    LayoutCachedWidth =14604
                    LayoutCachedHeight =975
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =12576
                    Top =480
                    Width =1320
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label878"
                    Caption ="Fee \015\012Amounts"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12576
                    LayoutCachedTop =480
                    LayoutCachedWidth =13896
                    LayoutCachedHeight =975
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =45
                    Width =705
                    FontSize =9
                    ForeColor =8355711
                    Name ="SchID"
                    ControlSource ="SchID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =45
                    LayoutCachedWidth =13275
                    LayoutCachedHeight =285
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =14616
                    Top =480
                    Width =720
                    Height =495
                    FontSize =8
                    BackColor =10092543
                    BorderColor =10092543
                    Name ="Label2206"
                    Caption ="LO \015\012%"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14616
                    LayoutCachedTop =480
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =975
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2112
                    Top =480
                    Width =1380
                    Height =495
                    Name ="EmptyCell2230"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2112
                    LayoutCachedTop =480
                    LayoutCachedWidth =3492
                    LayoutCachedHeight =975
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =2100
                    Top =480
                    Width =1380
                    Height =495
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label2232"
                    Caption ="Vendor Cost \015\012(per unit)"
                    FontName ="Segoe UI"
                    GroupTable =71
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2100
                    LayoutCachedTop =480
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =975
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =71
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9504
                    Top =480
                    Width =1080
                    Height =495
                    FontSize =9
                    TabIndex =3
                    ForeColor =10040879
                    Name ="Label16"
                    ControlSource ="=\"Payment (\" & [LeasePaidAdvArr] & \")\""
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9504
                    LayoutCachedTop =480
                    LayoutCachedWidth =10584
                    LayoutCachedHeight =975
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =2
                            Left =8604
                            Top =480
                            Width =840
                            Height =495
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label14"
                            Caption ="LRF"
                            FontName ="Segoe UI"
                            GroupTable =2
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =8604
                            LayoutCachedTop =480
                            LayoutCachedWidth =9444
                            LayoutCachedHeight =975
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =480
                    Width =1092
                    Height =495
                    FontSize =9
                    TabIndex =4
                    ForeColor =10040879
                    Name ="Label11"
                    ControlSource ="=[FundDateType] & \" Fund Date\""
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedTop =480
                    LayoutCachedWidth =7068
                    LayoutCachedHeight =975
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =2
                            Left =4920
                            Top =480
                            Width =996
                            Height =495
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label10"
                            Caption ="BLCD\015\012Date"
                            FontName ="Segoe UI"
                            GroupTable =2
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4920
                            LayoutCachedTop =480
                            LayoutCachedWidth =5916
                            LayoutCachedHeight =975
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =246
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =672
                    Width =1380
                    ColumnWidth =1530
                    FontSize =9
                    TabIndex =1
                    Name ="CostPerUnit"
                    ControlSource ="OECUnit"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =672
                    LayoutCachedWidth =2052
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Width =996
                    FontSize =9
                    TabIndex =4
                    Name ="BLCD"
                    ControlSource ="BLCD"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4920
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Width =1092
                    FontSize =9
                    TabIndex =5
                    Name ="EstFundDate"
                    ControlSource ="FundDateActive"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedWidth =7068
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7128
                    Width =780
                    ColumnWidth =1815
                    FontSize =9
                    TabIndex =6
                    Name ="MPY"
                    ControlSource ="MPYAllowed"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7128
                    LayoutCachedWidth =7908
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11892
                    Width =648
                    ColumnWidth =1815
                    FontSize =9
                    TabIndex =11
                    Name ="UnitRVPct"
                    ControlSource ="RVUnitPct"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11892
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12576
                    Width =1320
                    FontSize =9
                    TabIndex =12
                    Name ="FeePer"
                    ControlSource ="SyndFeeUnit"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12576
                    LayoutCachedWidth =13896
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9504
                    Width =1080
                    FontSize =9
                    TabIndex =9
                    Name ="LeasePymtUnit"
                    ControlSource ="[PaymentUnit]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9504
                    LayoutCachedWidth =10584
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7968
                    Width =576
                    FontSize =9
                    TabIndex =7
                    Name ="LeaseTermPrimary"
                    ControlSource ="LeaseTermPrimary"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7968
                    LayoutCachedWidth =8544
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8604
                    Width =840
                    FontSize =9
                    TabIndex =8
                    Name ="LRF"
                    ControlSource ="=Format([LRFUnit],\"0.00000%\")"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8604
                    LayoutCachedWidth =9444
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Width =1272
                    ColumnWidth =1605
                    FontSize =9
                    TabIndex =10
                    Name ="UnitResidual"
                    ControlSource ="RVUnit"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =11892
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3552
                    Width =1308
                    FontSize =9
                    TabIndex =3
                    Name ="UnitSalePrice"
                    ControlSource ="=([SalePrice]/[UnitsInGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3552
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Width =576
                    FontSize =9
                    ForeColor =10040879
                    Name ="Label1259"
                    ControlSource ="UnitCount"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =36
                    LayoutCachedWidth =612
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13932
                    Width =672
                    FontSize =9
                    TabIndex =13
                    ForeColor =8355711
                    Name ="Text2102"
                    ControlSource ="=[FeePer]/[CostPerUnit]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13932
                    LayoutCachedWidth =14604
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2112
                    Width =1380
                    FontSize =9
                    TabIndex =2
                    Name ="Text2221"
                    ControlSource ="OECVendorUnit"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2112
                    LayoutCachedWidth =3492
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin EmptyCell
                    Left =14616
                    Width =720
                    Name ="EmptyCell2277"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14616
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =662
                    Width =225
                    FontSize =9
                    TabIndex =14
                    ForeColor =10040879
                    Name ="Text2289"
                    ControlSource ="=\"@\""
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =662
                    LayoutCachedWidth =887
                    LayoutCachedHeight =240
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4440
            BreakLevel =2
            Name ="GroupFooter0"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =0
                    Left =4680
                    Top =420
                    Width =1282
                    Height =3600
                    Name ="Box2502"
                    LayoutCachedLeft =4680
                    LayoutCachedTop =420
                    LayoutCachedWidth =5962
                    LayoutCachedHeight =4020
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1995
                    Top =45
                    Width =10575
                    Height =270
                    FontSize =9
                    TabIndex =9
                    ForeColor =10040879
                    Name ="Label1153"
                    ControlSource ="=\"───────────────────────────────────────── Group \" & [GroupNum] & \" Totals  "
                        "─────────────────────────────────────────────────────►\""

                    LayoutCachedLeft =1995
                    LayoutCachedTop =45
                    LayoutCachedWidth =12570
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =10
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =570
                    Height =274
                    FontSize =9
                    FontWeight =700
                    BorderColor =10319446
                    Name ="UnitsInGroup"
                    ControlSource ="=Sum([UnitCount])"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =600
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =30
                    Width =1320
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    BorderColor =10040879
                    Name ="FeeGroupLine"
                    ControlSource ="SyndFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =12570
                    LayoutCachedTop =30
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =30
                    Width =645
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="FeePct"
                    ControlSource ="=[SyndFee]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =13920
                    LayoutCachedTop =30
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12585
                    Top =1005
                    Width =1305
                    FontSize =9
                    TabIndex =32
                    ForeColor =1643706
                    Name ="SyndPenaltyLine"
                    ControlSource ="=[TermPenaltyActual]"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="amount/expenses normally reducing the syndication fee - offent money used to sup"
                        "port Remarketing"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12585
                    LayoutCachedTop =1005
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =1245
                    RowStart =7
                    RowEnd =7
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10335
                            Top =1005
                            Width =2250
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label32"
                            Caption ="- Termination Penalty"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10335
                            LayoutCachedTop =1005
                            LayoutCachedWidth =12585
                            LayoutCachedHeight =1245
                            RowStart =7
                            RowEnd =7
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12585
                    Top =1350
                    Width =1305
                    FontSize =9
                    FontWeight =700
                    TabIndex =35
                    Name ="BankFeeLine"
                    ControlSource ="=[SyndFee]+Sum([StripRentActual])-Abs([TermPenaltyActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12585
                    LayoutCachedTop =1350
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =1590
                    RowStart =6
                    RowEnd =6
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =11325
                            Top =1350
                            Width =1260
                            Height =240
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label33"
                            Caption ="Net Synd Fee"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11325
                            LayoutCachedTop =1350
                            LayoutCachedWidth =12585
                            LayoutCachedHeight =1590
                            RowStart =6
                            RowEnd =6
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =1635
                    Width =1305
                    FontSize =9
                    TabIndex =39
                    Name ="UTALine"
                    ControlSource ="=[UTAGroup]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =1635
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =1875
                    RowStart =4
                    RowEnd =4
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =1620
                    Width =645
                    FontSize =8
                    TabIndex =36
                    Name ="UATPct"
                    ControlSource ="=IIf([UTAGroup]>0,[UTAGroup]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =1620
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =1920
                    Width =645
                    FontSize =8
                    TabIndex =42
                    Name ="PDPct"
                    ControlSource ="=IIf([PerdiemLine]>0,[PerdiemLine]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =1920
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =2160
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =2820
                    Width =1305
                    FontSize =9
                    FontWeight =700
                    TabIndex =54
                    Name ="TotalFeeLine"
                    ControlSource ="=[BankFeeLine]+[UTALine]+[PerdiemLine]+[ServFeeLine]-(Abs([InterimLine]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =3060
                    RowStart =8
                    RowEnd =8
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10215
                            Top =2820
                            Width =2280
                            Height =240
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label38"
                            Caption ="Group Sub Total"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10215
                            LayoutCachedTop =2820
                            LayoutCachedWidth =12495
                            LayoutCachedHeight =3060
                            RowStart =8
                            RowEnd =8
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =2805
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =52
                    Name ="SyndFeePct"
                    ControlSource ="=[TotalFeeLine]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =2805
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =3045
                    RowStart =8
                    RowEnd =8
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =1140
                    Width =612
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    Name ="Yield"
                    ControlSource ="IndexYield"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =1380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =1140
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label22"
                            Caption =" Yield"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =1140
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =1380
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =660
                    Width =1284
                    Height =480
                    FontSize =9
                    TabIndex =13
                    BorderColor =8355711
                    Name ="Text845"
                    ControlSource ="=\"<b>\" & [IndexYear] & \"</b> Year \" & [Index] & \"<br/>\" & [IndexDate]"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =4680
                    LayoutCachedTop =660
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =6
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin Rectangle
                    Left =10155
                    Top =3975
                    Width =5040
                    Height =15
                    Name ="Box880"
                    LayoutCachedLeft =10155
                    LayoutCachedTop =3975
                    LayoutCachedWidth =15195
                    LayoutCachedHeight =3990
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7650
                    Top =1635
                    Width =4860
                    FontSize =9
                    TabIndex =38
                    ForeColor =10040879
                    Name ="Text1262"
                    ControlSource ="MsgUTA"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =7650
                    LayoutCachedTop =1635
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =1875
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12585
                    Top =690
                    Width =1305
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =28
                    Name ="GrossFeeLine"
                    ControlSource ="=[SyndFee]+Sum([StripRentActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12585
                    LayoutCachedTop =690
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =975
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10335
                            Top =690
                            Width =2250
                            Height =240
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label1905"
                            Caption ="Synd Fee + Strip(s)"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10335
                            LayoutCachedTop =690
                            LayoutCachedWidth =12585
                            LayoutCachedHeight =930
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin Rectangle
                    Left =12623
                    Top =1290
                    Width =1267
                    Height =15
                    Name ="Box1480"
                    LayoutCachedLeft =12623
                    LayoutCachedTop =1290
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =1305
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =1860
                    Width =612
                    FontSize =9
                    TabIndex =21
                    Name ="Text1497"
                    ControlSource ="IndexBonus"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =2100
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =1860
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label1498"
                            Caption =" Bonus"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =1860
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =2100
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =990
                    Width =645
                    FontSize =8
                    TabIndex =30
                    Name ="Text1511"
                    ControlSource ="=IIf([SyndPenaltyLine]<>0,[SyndPenaltyLine]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =990
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =1230
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =5
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =675
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =26
                    Name ="Text1513"
                    ControlSource ="=[GrossFeeLine]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =675
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =915
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =1305
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =33
                    Name ="Text1514"
                    ControlSource ="=[BankFeeLine]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =1305
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =1545
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12585
                    Top =375
                    Width =1305
                    FontSize =9
                    TabIndex =12
                    Name ="StripsLine"
                    ControlSource ="=Sum([StripRentActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12585
                    LayoutCachedTop =375
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =615
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =360
                    Width =645
                    FontSize =8
                    TabIndex =11
                    Name ="Text1904"
                    ControlSource ="=IIf([StripsLine]>0,[StripsLine]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =360
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =600
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7680
                    Top =345
                    Width =4905
                    Height =285
                    FontSize =9
                    TabIndex =10
                    ForeColor =10040879
                    Name ="Label1903"
                    ControlSource ="MsgStrips"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =345
                    LayoutCachedWidth =12585
                    LayoutCachedHeight =630
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Left =12623
                    Top =645
                    Width =1267
                    Height =15
                    Name ="Box1906"
                    LayoutCachedLeft =12623
                    LayoutCachedTop =645
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =660
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12555
                    Top =2520
                    Width =1305
                    FontSize =9
                    TabIndex =50
                    Name ="ServFeeLine"
                    ControlSource ="=Sum([FSLFeePV])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12555
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =2760
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =2505
                    Width =645
                    FontSize =8
                    TabIndex =47
                    Name ="Text2061"
                    ControlSource ="=IIf([FSLFeePV]>0,[FSLFeePV]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =2505
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =2745
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7665
                    Top =2520
                    Width =4830
                    FontSize =9
                    TabIndex =49
                    ForeColor =10040879
                    Name ="txtServiceFeeDynamicLabel"
                    ControlSource ="MsgFSL"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =7665
                    LayoutCachedTop =2520
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2760
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12555
                    Top =2220
                    Width =1305
                    FontSize =9
                    TabIndex =46
                    Name ="InterimLine"
                    ControlSource ="=Sum([InterimExpenseGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12555
                    LayoutCachedTop =2220
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =2205
                    Width =645
                    FontSize =8
                    TabIndex =44
                    Name ="Text2064"
                    ControlSource ="=IIf([InterimLine]<>0,[InterimLine]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =2205
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =2445
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =3420
                    Width =1305
                    FontSize =9
                    TabIndex =59
                    Name ="HBResidual"
                    ControlSource ="HoldbackResidual"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =3420
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =3660
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10275
                            Top =3420
                            Width =2235
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2067"
                            Caption ="- Holdback Residual"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10275
                            LayoutCachedTop =3420
                            LayoutCachedWidth =12510
                            LayoutCachedHeight =3660
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =3405
                    Width =645
                    FontSize =8
                    TabIndex =58
                    Name ="Text2068"
                    ControlSource ="=IIf([HBResidual]>0,[HBResidual]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =3405
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =3645
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin Rectangle
                    Left =10155
                    Top =2775
                    Width =5040
                    Height =15
                    Name ="Box2070"
                    LayoutCachedLeft =10155
                    LayoutCachedTop =2775
                    LayoutCachedWidth =15195
                    LayoutCachedHeight =2790
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =3720
                    Width =1305
                    FontSize =9
                    TabIndex =61
                    Name ="HBConversion"
                    ControlSource ="HoldbackConversion"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =3720
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =3960
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10275
                            Top =3720
                            Width =2235
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2072"
                            Caption ="- Holdback Conversion"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10275
                            LayoutCachedTop =3720
                            LayoutCachedWidth =12510
                            LayoutCachedHeight =3960
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =3705
                    Width =645
                    FontSize =8
                    TabIndex =60
                    Name ="Text2073"
                    ControlSource ="=IIf([HBConversion]>0,[HBConversion]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =3705
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =3945
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12576
                    Top =4140
                    Width =1308
                    FontSize =9
                    FontWeight =700
                    TabIndex =68
                    Name ="FinalFee"
                    ControlSource ="=[TotalFeeLine]-([HBConversion]+[HBResidual]+[WitheldLine])+Nz([SyndMiscValue_1]"
                        ",0)+Nz([SyndMiscValue_2],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =12576
                    LayoutCachedTop =4140
                    LayoutCachedWidth =13884
                    LayoutCachedHeight =4380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =67
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =9000
                            Top =4140
                            Width =3510
                            Height =240
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label2075"
                            Caption ="Group Total Fee"
                            FontName ="Segoe UI"
                            GroupTable =67
                            LayoutCachedLeft =9000
                            LayoutCachedTop =4140
                            LayoutCachedWidth =12510
                            LayoutCachedHeight =4380
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =67
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13944
                    Top =4140
                    Width =648
                    FontSize =8
                    FontWeight =700
                    TabIndex =69
                    Name ="Text2076"
                    ControlSource ="=[FinalFee]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =13944
                    LayoutCachedTop =4140
                    LayoutCachedWidth =14592
                    LayoutCachedHeight =4380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =67
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =3120
                    Width =4140
                    Height =285
                    FontSize =9
                    TabIndex =56
                    ForeColor =10040879
                    Name ="Text2079"
                    ControlSource ="MsgWithHeld"
                    FontName ="Segoe UI"
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =6960
                    LayoutCachedTop =3120
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =3405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =3120
                    Width =1305
                    FontSize =9
                    TabIndex =57
                    Name ="WitheldLine"
                    ControlSource ="=[WithheldAmount]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =3120
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =3360
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =11160
                            Top =3120
                            Width =1350
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2081"
                            Caption ="- Advanced Rent"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11160
                            LayoutCachedTop =3120
                            LayoutCachedWidth =12510
                            LayoutCachedHeight =3360
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =3105
                    Width =645
                    FontSize =8
                    TabIndex =55
                    Name ="Text2082"
                    ControlSource ="=IIf([WitheldLine]>0,[WitheldLine]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =3105
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =3345
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =4080
                    Width =7080
                    FontSize =8
                    TabIndex =70
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="LastUpdate"
                    ControlSource ="=\"Last Modified by \" & [LastUpdatedBy] & \" on \" & [LastUpdatedDate]"
                    StatusBarText ="upadj"
                    FontName ="Segoe UI"
                    ControlTipText ="Rate Adjustment - If Upward"
                    ShowDatePicker =1

                    LayoutCachedTop =4080
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =4320
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3540
                    Top =30
                    Width =1305
                    Height =274
                    FontSize =9
                    TabIndex =3
                    BorderColor =10319446
                    Name ="Text2105"
                    ControlSource ="SalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =30
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =660
                    Top =30
                    Width =1380
                    Height =274
                    FontSize =9
                    TabIndex =1
                    BorderColor =10319446
                    Name ="OECLine"
                    ControlSource ="=Sum([OecGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =660
                    LayoutCachedTop =30
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9487
                    Top =30
                    Width =1073
                    Height =274
                    FontSize =9
                    TabIndex =4
                    BorderColor =10319446
                    Name ="Text2109"
                    ControlSource ="=Sum([PaymentGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =9487
                    LayoutCachedTop =30
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10592
                    Top =30
                    Width =1273
                    Height =274
                    FontSize =9
                    TabIndex =5
                    BorderColor =10319446
                    Name ="Text2111"
                    ControlSource ="=Sum([RVGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =10592
                    LayoutCachedTop =30
                    LayoutCachedWidth =11865
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12576
                    Top =4020
                    Width =1308
                    Height =0
                    FontSize =9
                    TabIndex =63
                    Name ="SyndMiscValue_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscValue1],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =12576
                    LayoutCachedTop =4020
                    LayoutCachedWidth =13884
                    LayoutCachedHeight =4020
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =67
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13944
                    Top =4020
                    Width =648
                    Height =0
                    FontSize =8
                    TabIndex =64
                    Name ="SyndMiscValue_1Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscValue1]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =13944
                    LayoutCachedTop =4020
                    LayoutCachedWidth =14592
                    LayoutCachedHeight =4020
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =67
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =4020
                    Width =3510
                    Height =0
                    FontSize =9
                    TabIndex =62
                    ForeColor =10040879
                    Name ="SyndMiscLabel_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscLabel1],Null)"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =9000
                    LayoutCachedTop =4020
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =4020
                    LayoutGroup =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =67
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12576
                    Top =4080
                    Width =1308
                    Height =0
                    FontSize =9
                    TabIndex =66
                    Name ="SyndMiscValue_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscValue2],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =12576
                    LayoutCachedTop =4080
                    LayoutCachedWidth =13884
                    LayoutCachedHeight =4080
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =67
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13944
                    Top =4080
                    Width =648
                    Height =0
                    FontSize =8
                    TabIndex =67
                    Name ="SyndMiscValue_2Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscValue2]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =13944
                    LayoutCachedTop =4080
                    LayoutCachedWidth =14592
                    LayoutCachedHeight =4080
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =67
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =4080
                    Width =3510
                    Height =0
                    FontSize =9
                    TabIndex =65
                    ForeColor =10040879
                    Name ="SyndMiscLabel_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscLabel2],Null)"
                    FontName ="Segoe UI"
                    GroupTable =67

                    LayoutCachedLeft =9000
                    LayoutCachedTop =4080
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =4080
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =67
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =720
                    Width =4560
                    Height =3300
                    FontSize =9
                    TabIndex =29
                    Name ="Text2171"
                    ControlSource ="SynMemo"

                    LayoutCachedLeft =60
                    LayoutCachedTop =720
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            Left =60
                            Top =480
                            Width =615
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2172"
                            Caption ="Notes:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =480
                            LayoutCachedWidth =675
                            LayoutCachedHeight =720
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7665
                    Top =2220
                    Width =4845
                    FontSize =9
                    TabIndex =45
                    ForeColor =10040879
                    Name ="InterminLabel"
                    ControlSource ="MsgInterim"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =7665
                    LayoutCachedTop =2220
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12570
                    Top =1920
                    Width =1305
                    FontSize =9
                    TabIndex =41
                    Name ="PerdiemLine"
                    ControlSource ="=Sum([PerdiemActive])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12570
                    LayoutCachedTop =1920
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =2160
                    RowStart =4
                    RowEnd =4
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7650
                    Top =1905
                    Width =4860
                    Height =285
                    FontSize =9
                    TabIndex =40
                    ForeColor =10040879
                    Name ="Text2183"
                    ControlSource ="=[MsgPerdiem]"
                    FontName ="Segoe UI"
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =7650
                    LayoutCachedTop =1905
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =2190
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =30
                    Width =645
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="Text2184"
                    ControlSource ="=[SyndFeeLO]/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =14565
                    LayoutCachedTop =30
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =990
                    Width =645
                    FontSize =8
                    TabIndex =31
                    Name ="Text2204"
                    ControlSource ="=IIf([TermPenaltyProjected]<>0,[TermPenaltyProjected]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =990
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =1230
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =5
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =675
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =27
                    Name ="Text2205"
                    ControlSource ="=([SyndFeeLO]+[StripsActual])/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =675
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =915
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =1305
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =34
                    Name ="Text2212"
                    ControlSource ="=(([SyndFeeLO]+[StripsActual])-Abs([TermPenaltyProjected]))/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =1305
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =1545
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =1620
                    Width =645
                    FontSize =8
                    TabIndex =37
                    Name ="Text2213"
                    ControlSource ="=IIf([UTAGroupLO]>0,[UTAGroupLO]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =1620
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =1920
                    Width =645
                    FontSize =8
                    TabIndex =43
                    Name ="Text2214"
                    ControlSource ="=IIf([PerDiemGroupLO]>0,[PerDiemGroupLO]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =1920
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =2160
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =2505
                    Width =645
                    FontSize =8
                    TabIndex =48
                    Name ="Text2215"
                    ControlSource ="=IIf([FSLFeePVLO]>0,[FSLFeePVLO]/[GroupCostLine],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =2505
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =2745
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14565
                    Top =2805
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =53
                    Name ="Text2216"
                    ControlSource ="=(([SyndFeeLO]+[StripsActual]+Nz([UTAGroupLO],0)+Nz([PerDiemGroupLO],0)+Nz([FSLF"
                        "eePVLO],0))-Abs([TermPenaltyProjected]))/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14565
                    LayoutCachedTop =2805
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =3045
                    RowStart =8
                    RowEnd =8
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14580
                    Top =4140
                    Width =645
                    FontSize =8
                    FontWeight =700
                    TabIndex =71
                    Name ="Text2217"
                    ControlSource ="=(([SyndFeeLO]+[StripsActual]+Nz([UTAGroupLO],0)+Nz([PerDiemGroupLO],0)+Nz([FSLF"
                        "eePVLO],0))-Abs([TermPenaltyProjected]))/[GroupCostLine]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =14580
                    LayoutCachedTop =4140
                    LayoutCachedWidth =15225
                    LayoutCachedHeight =4380
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2100
                    Top =30
                    Width =1380
                    Height =274
                    FontSize =9
                    TabIndex =2
                    BorderColor =10319446
                    Name ="GroupCostLine"
                    ControlSource ="=Sum([OECVendorGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2100
                    LayoutCachedTop =30
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =1140
                    Width =720
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="Text2238"
                    ControlSource ="IndexYieldLO"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =1380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =1860
                    Width =720
                    FontSize =9
                    TabIndex =22
                    Name ="Text2239"
                    ControlSource ="IndexBonusLO"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =2100
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =660
                    Width =1272
                    Height =480
                    FontSize =9
                    TabIndex =14
                    BorderColor =8355711
                    Name ="Text2252"
                    ControlSource ="=\"<b>\" & [IndexYearLO] & \"</b> Year \" & [IndexLO] & \"<br/>\" & [IndexDateLO"
                        "]"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =5976
                    LayoutCachedTop =660
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =6
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4680
                    Top =420
                    Width =1284
                    Height =240
                    FontSize =9
                    Name ="Label2253"
                    Caption ="SYND"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4680
                    LayoutCachedTop =420
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =660
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =74
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5976
                    Top =420
                    Width =1272
                    Height =240
                    FontSize =9
                    BackColor =10092543
                    ForeColor =10040879
                    Name ="Label2273"
                    Caption ="LO"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =420
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =660
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =74
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6030
                    Top =2805
                    Width =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =51
                    Name ="Text2291"
                    ControlSource ="IndexSpread"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6030
                    LayoutCachedTop =2805
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                End
                Begin EmptyCell
                    Left =7620
                    Top =3492
                    Width =4890
                    Name ="EmptyCell2325"
                    GroupTable =73
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7620
                    LayoutCachedTop =3492
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =3732
                    LayoutGroup =5
                    GroupTable =73
                End
                Begin EmptyCell
                    Left =7620
                    Top =3732
                    Width =4890
                    Name ="EmptyCell2326"
                    GroupTable =73
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7620
                    LayoutCachedTop =3732
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =3972
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    GroupTable =73
                End
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    Left =4680
                    Top =2100
                    Width =1284
                    Height =240
                    FontSize =9
                    Name ="Label2329"
                    Caption ="IRR-PT"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4680
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =2340
                    RowStart =6
                    RowEnd =6
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =74
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =2340
                    Width =612
                    FontSize =9
                    FontWeight =700
                    TabIndex =23
                    Name ="IndexSpreadIRR"
                    ControlSource ="IndexSpreadIRR"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =2580
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =2340
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2331"
                            Caption =" Spread"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =2340
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =2580
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =2580
                    Width =612
                    FontSize =9
                    TabIndex =24
                    Name ="IndexYieldIRR"
                    ControlSource ="IndexYieldIRR"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =2820
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =2580
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2333"
                            Caption =" Yield"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =2580
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =2820
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =1380
                    Width =612
                    FontSize =9
                    TabIndex =17
                    Name ="Text2369"
                    ControlSource ="IndexRate"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =1380
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2370"
                            Caption =" Index"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =1380
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =1620
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5352
                    Top =1620
                    Width =612
                    FontSize =9
                    FontWeight =700
                    TabIndex =19
                    Name ="Text2371"
                    ControlSource ="IndexSpread"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =4680
                            Top =1620
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2372"
                            Caption =" Spread"
                            FontName ="Segoe UI"
                            GroupTable =74
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4680
                            LayoutCachedTop =1620
                            LayoutCachedWidth =5355
                            LayoutCachedHeight =1860
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =6
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =74
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =1380
                    Width =720
                    FontSize =9
                    TabIndex =18
                    Name ="Text2375"
                    ControlSource ="=[IndexRateLO]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5976
                    Top =1620
                    Width =720
                    FontSize =9
                    FontWeight =700
                    TabIndex =20
                    Name ="Text2376"
                    ControlSource ="=[IndexSpreadLO]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5976
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =1140
                    Width =552
                    Name ="EmptyCell2397"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =1140
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =1380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =1860
                    Width =552
                    Name ="EmptyCell2402"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =2100
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =5976
                    Top =2100
                    Width =720
                    Name ="EmptyCell2470"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =2340
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =2100
                    Width =552
                    Name ="EmptyCell2471"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =2340
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =5976
                    Top =2340
                    Width =720
                    Name ="EmptyCell2476"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =2580
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =2340
                    Width =552
                    Name ="EmptyCell2477"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =2580
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =5976
                    Top =2580
                    Width =720
                    Name ="EmptyCell2484"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =2820
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =2580
                    Width =552
                    Name ="EmptyCell2485"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =2580
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =2820
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin Rectangle
                    BackStyle =0
                    Left =5962
                    Top =420
                    Width =1282
                    Height =1728
                    Name ="Box2503"
                    LayoutCachedLeft =5962
                    LayoutCachedTop =420
                    LayoutCachedWidth =7244
                    LayoutCachedHeight =2148
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                End
                Begin EmptyCell
                    Left =6696
                    Top =1380
                    Width =552
                    Name ="EmptyCell2528"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =1620
                    Width =552
                    Name ="EmptyCell2535"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =1620
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    Left =4680
                    Top =2820
                    Width =1284
                    Height =240
                    FontSize =9
                    Name ="Label2536"
                    Caption ="RV Guar"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4680
                    LayoutCachedTop =2820
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =3060
                    RowStart =9
                    RowEnd =9
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =0
                    BackTint =95.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =74
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =3060
                    Width =1284
                    FontSize =9
                    TabIndex =25
                    Name ="RVGuar"
                    ControlSource ="RVGuar"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4680
                    LayoutCachedTop =3060
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =3300
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    LayoutGroup =6
                    ThemeFontIndex =-1
                    GroupTable =74
                    CurrencySymbol ="$"
                End
                Begin EmptyCell
                    Left =5976
                    Top =2820
                    Width =720
                    Name ="EmptyCell2549"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =3060
                    RowStart =9
                    RowEnd =9
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =2820
                    Width =552
                    Name ="EmptyCell2550"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =3060
                    RowStart =9
                    RowEnd =9
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =5976
                    Top =3060
                    Width =720
                    Name ="EmptyCell2554"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5976
                    LayoutCachedTop =3060
                    LayoutCachedWidth =6696
                    LayoutCachedHeight =3300
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GroupTable =74
                End
                Begin EmptyCell
                    Left =6696
                    Top =3060
                    Width =552
                    Name ="EmptyCell2555"
                    GroupTable =74
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6696
                    LayoutCachedTop =3060
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =3300
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =74
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1379
            BreakLevel =1
            Name ="GroupFooter3"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Width =15300
                    Height =1379
                    Name ="ScheduleEconomicsSyndReportSub"
                    SourceObject ="Report.ScheduleEconomicsSyndReportSub"
                    LinkChildFields ="schid"
                    LinkMasterFields ="schid"

                    LayoutCachedWidth =15300
                    LayoutCachedHeight =1379
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =120
                    Width =3240
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =0
                    Name ="Text1510"
                    Caption ="Fleet Advantage Confidential"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4560
                    Width =10710
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =8355711
                    Name ="Text1491"
                    ControlSource ="=[Client] & \" - Schedule \" & [Schedule] & \" Final Economics Ver 2.1\""

                    LayoutCachedLeft =4560
                    LayoutCachedWidth =15270
                    LayoutCachedHeight =225
                    ForeTint =50.0
                End
            End
        End
    End
End
CodeBehindForm
' See "ScheduleEconomicsSyndReport.cls"
