Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =25560
    DatasheetFontHeight =11
    ItemSuffix =220
    Left =6255
    Top =6060
    Right =21735
    Bottom =22845
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x11830e1fd7cbe440
    End
    RecordSource ="SELECT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null  And InvoiceP"
        "aid Between #1/1/2015# AND #12/31/2015# Order By SourceCompany;"
    Caption ="MGR"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
        Begin ListBox
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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =3195
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =1
                    PictureType =2
                    Top =3000
                    Width =25560
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =3000
                    LayoutCachedWidth =25560
                    LayoutCachedHeight =3195
                    TabIndex =39
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Width =25560
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =25560
                    LayoutCachedHeight =540
                End
                Begin Line
                    OverlapFlags =93
                    Left =21477
                    Top =885
                    Width =0
                    Height =792
                    BorderColor =12349952
                    Name ="Line203"
                    LayoutCachedLeft =21477
                    LayoutCachedTop =885
                    LayoutCachedWidth =21477
                    LayoutCachedHeight =1677
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =20205
                    Top =1680
                    Width =2685
                    Height =270
                    FontSize =8
                    ForeColor =12349952
                    Name ="Label44"
                    Caption ="┌───── Quick Set Date Rages ─────┐"
                    LayoutCachedLeft =20205
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22890
                    LayoutCachedHeight =1950
                    BackThemeColorIndex =1
                End
                Begin Line
                    OverlapFlags =95
                    Left =600
                    Top =885
                    Width =20880
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =600
                    LayoutCachedTop =885
                    LayoutCachedWidth =21480
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =3360
                    Top =2790
                    Width =705
                    Height =270
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="♦ Asset"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3360
                    LayoutCachedTop =2790
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =3060
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =12555
                    Top =2280
                    Width =2955
                    Height =270
                    FontSize =8
                    Name ="Label20"
                    Caption ="┌───────────── Cost ─────────────┐"
                    LayoutCachedLeft =12555
                    LayoutCachedTop =2280
                    LayoutCachedWidth =15510
                    LayoutCachedHeight =2550
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =2655
                    Top =2790
                    Width =645
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Year"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2655
                    LayoutCachedTop =2790
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =3060
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =5580
                    Top =2790
                    Width =1380
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Make"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5580
                    LayoutCachedTop =2790
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =3060
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =7020
                    Top =2790
                    Width =1597
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label4"
                    Caption ="♦ Buyer"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7020
                    LayoutCachedTop =2790
                    LayoutCachedWidth =8617
                    LayoutCachedHeight =3060
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =11190
                    Top =2790
                    Width =630
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Days In"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =11190
                    LayoutCachedTop =2790
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =3060
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =12045
                    Top =2790
                    Width =1245
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label9"
                    Caption ="Buy Out"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12045
                    LayoutCachedTop =2790
                    LayoutCachedWidth =13290
                    LayoutCachedHeight =3060
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =13350
                    Top =2790
                    Width =1245
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="Expenses"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13350
                    LayoutCachedTop =2790
                    LayoutCachedWidth =14595
                    LayoutCachedHeight =3060
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =16080
                    Top =2790
                    Width =1245
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Sold Price"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16080
                    LayoutCachedTop =2790
                    LayoutCachedWidth =17325
                    LayoutCachedHeight =3060
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =17355
                    Top =2790
                    Width =1245
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BackColor =12750136
                    Name ="Label15"
                    Caption ="Profit"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17355
                    LayoutCachedTop =2790
                    LayoutCachedWidth =18600
                    LayoutCachedHeight =3060
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =18660
                    Top =2790
                    Width =1005
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Margin"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18660
                    LayoutCachedTop =2790
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =3060
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16140
                    Top =2475
                    Width =1245
                    Height =270
                    ColumnOrder =21
                    FontSize =10
                    FontWeight =700
                    TabIndex =29
                    ForeColor =10040879
                    Name ="SoldPriceLine"
                    ControlSource ="=Sum([UnitSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =16140
                    LayoutCachedTop =2475
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =2745
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12075
                    Top =2475
                    Width =1245
                    Height =270
                    ColumnOrder =25
                    FontSize =10
                    FontWeight =700
                    TabIndex =26
                    ForeColor =10040879
                    Name ="BuyOutLine"
                    ControlSource ="=Sum([BuyOut])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =12075
                    LayoutCachedTop =2475
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =2745
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Top =2475
                    Width =1245
                    Height =270
                    ColumnOrder =24
                    FontSize =10
                    FontWeight =700
                    TabIndex =27
                    ForeColor =10040879
                    Name ="TotalExpenseLine"
                    ControlSource ="=Sum([TotalExpense])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =13380
                    LayoutCachedTop =2475
                    LayoutCachedWidth =14625
                    LayoutCachedHeight =2745
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17445
                    Top =2475
                    Width =1245
                    Height =270
                    ColumnOrder =23
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    ForeColor =10040879
                    Name ="TotalProfitLine"
                    ControlSource ="=[SoldPriceLine]-[TotalCostLine]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =17445
                    LayoutCachedTop =2475
                    LayoutCachedWidth =18690
                    LayoutCachedHeight =2745
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18750
                    Top =2475
                    Width =840
                    Height =270
                    ColumnOrder =22
                    FontSize =9
                    FontWeight =700
                    TabIndex =31
                    ForeColor =10040879
                    Name ="YearMargin"
                    ControlSource ="=IIf([TotalProfitLine]=0,0,([TotalProfitLine]/[TotalCostLine]))"
                    Format ="Percent"
                    ShowDatePicker =0

                    LayoutCachedLeft =18750
                    LayoutCachedTop =2475
                    LayoutCachedWidth =19590
                    LayoutCachedHeight =2745
                    BorderThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =135
                    Top =2790
                    Width =2460
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="♦ Source Company"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =135
                    LayoutCachedTop =2790
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =3060
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =8670
                    Top =2790
                    Width =1155
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Status"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =8670
                    LayoutCachedTop =2790
                    LayoutCachedWidth =9825
                    LayoutCachedHeight =3060
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =10185
                    Top =2340
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =25
                    Name ="ResetForm"
                    Caption ="Reset Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10185
                    LayoutCachedTop =2340
                    LayoutCachedWidth =11738
                    LayoutCachedHeight =2580
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2524406
                    BorderColor =2524406
                    HoverThemeColorIndex =9
                    HoverShade =50.0
                    PressedColor =2524406
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =20940
                    Top =2790
                    Width =1155
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label35"
                    Caption ="Invoice Sent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20940
                    LayoutCachedTop =2790
                    LayoutCachedWidth =22095
                    LayoutCachedHeight =3060
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =19725
                    Top =2790
                    Width =1155
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label36"
                    Caption ="BuyOutDate"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19725
                    LayoutCachedTop =2790
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =3060
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =22155
                    Top =2790
                    Width =1155
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Invoice Paid"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =22155
                    LayoutCachedTop =2790
                    LayoutCachedWidth =23310
                    LayoutCachedHeight =3060
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21000
                    Top =2490
                    Width =1155
                    Height =270
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =33
                    BorderColor =4210752
                    Name ="InvoiceSentEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =21000
                    LayoutCachedTop =2490
                    LayoutCachedWidth =22155
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21000
                    Top =2160
                    Width =1155
                    Height =270
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =23
                    BorderColor =4210752
                    Name ="InvoiceSentStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =21000
                    LayoutCachedTop =2160
                    LayoutCachedWidth =22155
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =22215
                    Top =2490
                    Width =1155
                    Height =270
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =34
                    BorderColor =4210752
                    Name ="SoldDateEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =22215
                    LayoutCachedTop =2490
                    LayoutCachedWidth =23370
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =22215
                    Top =2160
                    Width =1155
                    Height =270
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =24
                    BorderColor =4210752
                    Name ="SoldDateStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =22215
                    LayoutCachedTop =2160
                    LayoutCachedWidth =23370
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19770
                    Top =2490
                    Width =1155
                    Height =270
                    ColumnOrder =15
                    FontSize =10
                    TabIndex =32
                    BorderColor =4210752
                    Name ="BuyOutEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =19770
                    LayoutCachedTop =2490
                    LayoutCachedWidth =20925
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19770
                    Top =2160
                    Width =1155
                    Height =270
                    ColumnOrder =16
                    FontSize =10
                    TabIndex =22
                    BorderColor =4210752
                    Name ="BuyOutStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =19770
                    LayoutCachedTop =2160
                    LayoutCachedWidth =20925
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10380
                    Width =29
                    Height =270
                    ColumnOrder =17
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =10380
                    LayoutCachedWidth =10409
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =9360
                    Top =240
                    Width =765
                    Height =210
                    ColumnOrder =18
                    FontSize =8
                    TabIndex =1
                    Name ="txtCurrent"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =240
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =450
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11833
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =19
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitsShown"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =11833
                    LayoutCachedTop =240
                    LayoutCachedWidth =12522
                    LayoutCachedHeight =540
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =11808
                            Top =60
                            Width =750
                            Height =540
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            ForeColor =8872517
                            Name ="Label63"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11808
                            LayoutCachedTop =60
                            LayoutCachedWidth =12558
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =14655
                    Top =2790
                    Width =1245
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label61"
                    Caption ="Total Cost"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14655
                    LayoutCachedTop =2790
                    LayoutCachedWidth =15900
                    LayoutCachedHeight =3060
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14685
                    Top =2475
                    Width =1245
                    Height =270
                    ColumnOrder =20
                    FontSize =10
                    FontWeight =700
                    TabIndex =28
                    ForeColor =10040879
                    Name ="TotalCostLine"
                    ControlSource ="=[BuyOutLine]+[TotalExpenseLine]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =14685
                    LayoutCachedTop =2475
                    LayoutCachedWidth =15930
                    LayoutCachedHeight =2745
                    RowStart =1
                    RowEnd =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =4125
                    Top =2790
                    Width =1395
                    Height =270
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label70"
                    Caption ="Unit Type"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4125
                    LayoutCachedTop =2790
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =3060
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =120
                    Top =960
                    Width =2460
                    Height =1800
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstSourceCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SourceCompany, ClientId From RemarketingReportAllUnits WHERE [UNITId] is "
                        "not null  And InvoicePaid Between #1/1/2015# AND #12/31/2015#  GROUP BY SourceCo"
                        "mpany, ClientId ORDER BY SourceCompany;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2640
                    Top =960
                    Width =705
                    Height =1800
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModelYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitYr From RemarketingReportAllUnits WHERE [UNITId] is not null  And Inv"
                        "oicePaid Between #1/1/2015# AND #12/31/2015#  GROUP BY UnitYr ORDER BY UnitYr de"
                        "sc;"
                    ColumnWidths ="1440;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2640
                    LayoutCachedTop =960
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4110
                    Top =960
                    Width =1410
                    Height =1800
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstUnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitType From RemarketingReportAllUnits WHERE [UNITId] is not null  And I"
                        "nvoicePaid Between #1/1/2015# AND #12/31/2015#  GROUP BY UnitType ORDER BY UnitT"
                        "ype;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =4110
                    LayoutCachedTop =960
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5580
                    Top =960
                    Width =1380
                    Height =1800
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMake From RemarketingReportAllUnits WHERE [UNITId] is not null  And I"
                        "nvoicePaid Between #1/1/2015# AND #12/31/2015#  GROUP BY UnitMake ORDER BY UnitM"
                        "ake;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =5580
                    LayoutCachedTop =960
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7020
                    Top =960
                    Width =1590
                    Height =1800
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBuyer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BuyerName,BuyerId From RemarketingReportAllUnits WHERE [UNITId] is not nu"
                        "ll  And InvoicePaid Between #1/1/2015# AND #12/31/2015#  GROUP BY BuyerName, Buy"
                        "erId ORDER BY BuyerName;"
                    ColumnWidths ="144;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =7020
                    LayoutCachedTop =960
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10197
                    Top =1080
                    Width =1553
                    Height =1140
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="cmdSetFilters"
                    Caption ="After Setting Filters Click HERE\015\012Run the Results"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10197
                    LayoutCachedTop =1080
                    LayoutCachedWidth =11750
                    LayoutCachedHeight =2220
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    BorderColor =12750136
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8670
                    Top =960
                    Width =1382
                    Height =1800
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.UnitStatus, [UnitStatusDesc] & \" - \" & [UnitStatus] A"
                        "S [Desc] FROM LookUp_UnitStatus WHERE (((LookUp_UnitStatus.[UnitStatus]) In (SEL"
                        "ECT UnitStatus FROM RemarketingReportAllUnits GROUP BY UnitStatus ))) ORDER BY ["
                        "UnitStatusDesc] & \" - \" & [UnitStatus];"
                    ColumnWidths ="0;944"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =8670
                    LayoutCachedTop =960
                    LayoutCachedWidth =10052
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =120
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearClients"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =120
                    LayoutCachedTop =780
                    LayoutCachedWidth =660
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =2640
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearYear"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =780
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =4110
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearUnitType"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =4110
                    LayoutCachedTop =780
                    LayoutCachedWidth =4650
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =5580
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearMake"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =5580
                    LayoutCachedTop =780
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =7020
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBuyer"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =7020
                    LayoutCachedTop =780
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =8670
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearStatus"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =8670
                    LayoutCachedTop =780
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =1005
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10197
                    Top =780
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ClearAndRun"
                    Caption ="Clear Filters and Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10197
                    LayoutCachedTop =780
                    LayoutCachedWidth =11750
                    LayoutCachedHeight =1020
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    BorderColor =12750136
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =16980
                    Top =780
                    Width =2355
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearAll"
                    Caption ="◄ Click to clear All Selections  ►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =16980
                    LayoutCachedTop =780
                    LayoutCachedWidth =19335
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9900
                    Top =2790
                    Width =1215
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="Odometer"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =9900
                    LayoutCachedTop =2790
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =3060
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11895
                    Top =2790
                    Width =90
                    Height =270
                    Name ="EmptyCell149"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11895
                    LayoutCachedTop =2790
                    LayoutCachedWidth =11985
                    LayoutCachedHeight =3060
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15960
                    Top =2790
                    Width =90
                    Height =270
                    Name ="EmptyCell153"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15960
                    LayoutCachedTop =2790
                    LayoutCachedWidth =16050
                    LayoutCachedHeight =3060
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =23430
                    Top =2205
                    Width =390
                    Height =210
                    FontSize =8
                    Name ="Label159"
                    Caption ="Start"
                    LayoutCachedLeft =23430
                    LayoutCachedTop =2205
                    LayoutCachedWidth =23820
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =23430
                    Top =2535
                    Width =390
                    Height =210
                    FontSize =8
                    Name ="Label160"
                    Caption ="End"
                    LayoutCachedLeft =23430
                    LayoutCachedTop =2535
                    LayoutCachedWidth =23820
                    LayoutCachedHeight =2745
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =16335
                    Top =2280
                    Width =3195
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌──────────── Income ────────────┐"
                    LayoutCachedLeft =16335
                    LayoutCachedTop =2280
                    LayoutCachedWidth =19530
                    LayoutCachedHeight =2550
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =22215
                    Top =1905
                    Width =382
                    Height =225
                    FontSize =7
                    TabIndex =19
                    ForeColor =12349952
                    Name ="Sold0"
                    Caption ="2016"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =22215
                    LayoutCachedTop =1905
                    LayoutCachedWidth =22597
                    LayoutCachedHeight =2130
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =22605
                    Top =1905
                    Width =382
                    Height =225
                    FontSize =7
                    TabIndex =20
                    ForeColor =12349952
                    Name ="Sold1"
                    Caption ="2015"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =22605
                    LayoutCachedTop =1905
                    LayoutCachedWidth =22987
                    LayoutCachedHeight =2130
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =21000
                    Top =1905
                    Width =375
                    Height =225
                    FontSize =7
                    TabIndex =16
                    ForeColor =12349952
                    Name ="Sent0"
                    Caption ="2016"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =21000
                    LayoutCachedTop =1905
                    LayoutCachedWidth =21375
                    LayoutCachedHeight =2130
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =21390
                    Top =1905
                    Width =375
                    Height =225
                    FontSize =7
                    TabIndex =17
                    ForeColor =12349952
                    Name ="Sent1"
                    Caption ="2015"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =21390
                    LayoutCachedTop =1905
                    LayoutCachedWidth =21765
                    LayoutCachedHeight =2130
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =19785
                    Top =1905
                    Width =382
                    Height =225
                    FontSize =7
                    TabIndex =13
                    ForeColor =12349952
                    Name ="Buy0"
                    Caption ="2016"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =19785
                    LayoutCachedTop =1905
                    LayoutCachedWidth =20167
                    LayoutCachedHeight =2130
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =20175
                    Top =1905
                    Width =375
                    Height =225
                    FontSize =7
                    TabIndex =14
                    ForeColor =12349952
                    Name ="Buy1"
                    Caption ="2015"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =20175
                    LayoutCachedTop =1905
                    LayoutCachedWidth =20550
                    LayoutCachedHeight =2130
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =22995
                    Top =1905
                    Width =382
                    Height =225
                    FontSize =7
                    TabIndex =21
                    ForeColor =12349952
                    Name ="Sold2"
                    Caption ="2014"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =22995
                    LayoutCachedTop =1905
                    LayoutCachedWidth =23377
                    LayoutCachedHeight =2130
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =20565
                    Top =1905
                    Width =375
                    Height =225
                    FontSize =7
                    TabIndex =15
                    ForeColor =12349952
                    Name ="Buy2"
                    Caption ="2014"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =20565
                    LayoutCachedTop =1905
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =2130
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =21780
                    Top =1905
                    Width =375
                    Height =225
                    FontSize =7
                    TabIndex =18
                    ForeColor =12349952
                    Name ="Sent2"
                    Caption ="2014"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =21780
                    LayoutCachedTop =1905
                    LayoutCachedWidth =22155
                    LayoutCachedHeight =2130
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =16425
                    Top =1140
                    Width =2280
                    Height =1080
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =16425
                    LayoutCachedTop =1140
                    LayoutCachedWidth =18705
                    LayoutCachedHeight =2220
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =16140
                    Top =1500
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =16140
                    LayoutCachedTop =1500
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =1920
                    TabIndex =38
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin ListBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =16664
                    Top =1140
                    Width =2040
                    Height =1080
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BackColor =14277081
                    ForeColor =0
                    BorderColor =4210752
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags, LaunchReportsList.FormGroup FROM LaunchReportsList WH"
                        "ERE (((LaunchReportsList.SecurityLevel)=10) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashboard\")) ORDER BY LaunchReportsList.ReportPublicName, LaunchReportsList"
                        ".ListOrder;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =16664
                    LayoutCachedTop =1140
                    LayoutCachedWidth =18704
                    LayoutCachedHeight =2220
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13290
                    Top =960
                    Height =1320
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="TargetValueList"
                    RowSourceType ="Value List"
                    RowSource ="0;All;1;\"Buy Out\";2;\"No Buy Out\";3;\"Expenses\";4;\"No Expenses\""
                    ColumnWidths ="0;944"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"
                    ControlTipText ="Select the type of Unit Cost to filter by, Results will run on click, single sel"
                        "ection"
                    GridlineColor =0

                    LayoutCachedLeft =13290
                    LayoutCachedTop =960
                    LayoutCachedWidth =14730
                    LayoutCachedHeight =2280
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =13290
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearTarget"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =13290
                    LayoutCachedTop =780
                    LayoutCachedWidth =13830
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =223
                    Top =60
                    Width =4260
                    Height =555
                    FontSize =20
                    ForeColor =8872517
                    Name ="Label31"
                    Caption ="Accounting Dashboard"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =615
                End
                Begin Label
                    OverlapFlags =215
                    Left =30
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label294"
                    Caption ="Remarketing"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =30
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =240
                End
                Begin CommandButton
                    OverlapFlags =223
                    PictureType =1
                    Left =23115
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =35
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =23115
                    LayoutCachedTop =180
                    LayoutCachedWidth =23865
                    LayoutCachedHeight =408
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =22560
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =36
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =22560
                    LayoutCachedTop =180
                    LayoutCachedWidth =23115
                    LayoutCachedHeight =408
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =24060
                    Top =960
                    Width =1382
                    Height =1800
                    FontSize =8
                    FontWeight =700
                    TabIndex =37
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.BankPartner, RemarketingReportAllUnits.BankId F"
                        "ROM RemarketingReportAllUnits GROUP BY RemarketingReportAllUnits.BankPartner, Re"
                        "marketingReportAllUnits.BankId ORDER BY RemarketingReportAllUnits.BankPartner;"
                    ColumnWidths ="1440;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =24060
                    LayoutCachedTop =960
                    LayoutCachedWidth =25442
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =24060
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBanks"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =24060
                    LayoutCachedTop =780
                    LayoutCachedWidth =24600
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =24135
                    Top =2790
                    Width =1245
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label219"
                    Caption ="Assignee"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =24135
                    LayoutCachedTop =2790
                    LayoutCachedWidth =25380
                    LayoutCachedHeight =3075
                    LayoutGroup =6
                    GroupTable =8
                End
            End
        End
        Begin Section
            Height =299
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11190
                    Width =630
                    Height =255
                    FontSize =10
                    TabIndex =11
                    BorderColor =4210752
                    Name ="DaysInInventory"
                    ControlSource ="[DaysInInventory]"
                    StatusBarText ="Short Name for Reporting"
                    ControlTipText ="Days after 'Parked' in INV (before Sold Date, if sold). A neg value reflect days"
                        " until parked."
                    ConditionalFormat = Begin
                        0x010000007e000000030000000000000000000000000000000200000001000000 ,
                        0x00000000d7edda00000000000000000005000000080000000100000000000000 ,
                        0xffffcc0000000000060000000b0000000e00000001000000ffffff0099330000 ,
                        0x300000003400350000003400360000003700350000003700360000000000
                    End
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =11190
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x01000300000000000000000000000100000000000000d7edda00010000003000 ,
                        0x0200000034003500000000000000000000000000000000000000000000000000 ,
                        0x000100000000000000ffffcc0002000000340036000200000037003500000000 ,
                        0x0000000000000000000000000000000000000600000001000000ffffff009933 ,
                        0x00000200000037003600000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4125
                    Width =1395
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BorderColor =0
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    StatusBarText ="unit make"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =4125
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14655
                    Width =1245
                    Height =255
                    FontSize =10
                    TabIndex =14
                    BorderColor =0
                    Name ="UnitCost"
                    ControlSource ="=[TotalExpense]+[BuyOut]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =14655
                    LayoutCachedWidth =15900
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3360
                    Width =705
                    Height =255
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =5
                    BorderColor =0
                    ForeColor =4754549
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2655
                    Width =645
                    Height =255
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =4
                    BorderColor =0
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =2655
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5580
                    Width =1380
                    Height =255
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    StatusBarText ="unit make"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =5580
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12045
                    Width =1245
                    Height =255
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="BuyOut"
                    ControlSource ="BuyOut"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =12045
                    LayoutCachedWidth =13290
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16080
                    Width =1245
                    Height =255
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="UnitSold"
                    ControlSource ="UnitSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =16080
                    LayoutCachedWidth =17325
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Width =1597
                    Height =255
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =8
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="Buyer"
                    ControlSource ="Buyer"
                    StatusBarText ="Short Name for Reporting"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =7020
                    LayoutCachedWidth =8617
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13350
                    Width =1245
                    Height =255
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="TotalExpense"
                    ControlSource ="TotalExpense"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13350
                    LayoutCachedWidth =14595
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17355
                    Width =1245
                    Height =255
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =16
                    BorderColor =0
                    Name ="Profit"
                    ControlSource ="TotalProfit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =17355
                    LayoutCachedWidth =18600
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18660
                    Width =1005
                    Height =255
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="Margin"
                    ControlSource ="Margin"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =18660
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =135
                    Width =2460
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="SourceCompany"
                    ControlSource ="SourceCompany"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =135
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20940
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =19
                    BorderColor =0
                    Name ="InvoiceSent"
                    ControlSource ="InvoiceSent"
                    Format ="Short Date"
                    StatusBarText ="date unit parked"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =20940
                    LayoutCachedWidth =22095
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19725
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =18
                    BorderColor =0
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    Format ="Short Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19725
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =22155
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =20
                    BorderColor =0
                    Name ="InvoicePaid"
                    ControlSource ="InvoicePaid"
                    Format ="Short Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =22155
                    LayoutCachedWidth =23310
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =15
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =15
                    LayoutCachedTop =30
                    LayoutCachedWidth =60
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =75
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =1
                    Name ="BuyerID"
                    ControlSource ="BuyerID"

                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =120
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8670
                    Width =1155
                    Height =255
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =9
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =8670
                    LayoutCachedWidth =9825
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9900
                    Width =1215
                    Height =255
                    FontSize =10
                    TabIndex =10
                    BorderColor =0
                    Name ="Odometer"
                    ControlSource ="Odometer"
                    Format ="Standard"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =9900
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11895
                    Width =90
                    Height =255
                    Name ="EmptyCell150"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11895
                    LayoutCachedWidth =11985
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15960
                    Width =90
                    Height =255
                    Name ="EmptyCell154"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15960
                    LayoutCachedWidth =16050
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =240
                    Top =255
                    Width =25200
                    Height =44
                    TabIndex =2
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =255
                    LayoutCachedWidth =25440
                    LayoutCachedHeight =299
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =24135
                    Width =1245
                    Height =255
                    FontSize =10
                    TabIndex =21
                    BorderColor =0
                    Name ="Text216"
                    ControlSource ="BankPartner"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =24135
                    LayoutCachedWidth =25380
                    LayoutCachedHeight =255
                    LayoutGroup =5
                    GroupTable =7
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "RemarketingDashboardAccounting.cls"
