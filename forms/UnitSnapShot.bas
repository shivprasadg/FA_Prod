Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =238
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7725
    DatasheetFontHeight =11
    ItemSuffix =150
    Left =6060
    Top =3990
    Right =13785
    Bottom =6645
    Filter ="UnitID = 281283"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xd7bb66bc8b39e540
    End
    RecordSource ="SELECT vw_SixKeys.ClientShortName, vw_SixKeys.MLNo, vw_SixKeys.Schedule, vw_SixK"
        "eys.UnitGroup, Units.* FROM Units INNER JOIN vw_SixKeys ON Units.UnitID = vw_Six"
        "Keys.UnitId;"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =405
            BackColor =15527148
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =45
                    Width =7725
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Text208"
                    ControlSource ="=\"<B>\" & UCase([ClientShortName]) & \"</b> Lease No <b>\" & [mlno] & \" </b>Sc"
                        "h<b> \" & [Schedule] & \" </b>Grp<b>  \" & [UnitGroup] & \"</b>\""
                    TextFormat =1

                    LayoutCachedTop =45
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =405
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =2
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =2272
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =75
                    Width =4999
                    FontSize =10
                    BackColor =15527148
                    BorderColor =0
                    Name ="UnitDesc"
                    ControlSource ="=[unityr] & \" \" & [UnitMake] & \" \" & [UnitModel]"

                    LayoutCachedLeft =75
                    LayoutCachedWidth =5074
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =1440
                    Width =945
                    Height =225
                    FontSize =9
                    TabIndex =10
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483617
                    Name ="docSchedSent"
                    ControlSource ="UnitCOASent"
                    Format ="Short Date"
                    InputMask ="\"Sched Sent\""
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =1440
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1665
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =105
                            Top =1440
                            Width =990
                            Height =225
                            FontSize =9
                            Name ="Label15"
                            Caption ="COA sent"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1665
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =1725
                    Width =945
                    Height =225
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    BorderColor =-2147483617
                    ForeColor =-2147483617
                    Name ="docSchedBack"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =1725
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =105
                            Top =1725
                            Width =990
                            Height =225
                            FontSize =8
                            Name ="Label16"
                            Caption ="Accepted:"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =1725
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1950
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4590
                    Top =270
                    Width =1200
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =14
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    RightPadding =38

                    LayoutCachedLeft =4590
                    LayoutCachedTop =270
                    LayoutCachedWidth =5790
                    LayoutCachedHeight =525
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4005
                            Top =270
                            Width =525
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label9"
                            Caption ="Type:"
                            LayoutCachedLeft =4005
                            LayoutCachedTop =270
                            LayoutCachedWidth =4530
                            LayoutCachedHeight =525
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =585
                    Width =945
                    Height =225
                    FontSize =9
                    TabIndex =4
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483617
                    Name ="EstDelvDate"
                    ControlSource ="UnitEstDelDate"
                    StatusBarText ="schedule number starts with 1"
                    ControlTipText ="Double click to open schedule"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =585
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =105
                            Top =585
                            Width =990
                            Height =225
                            FontSize =9
                            Name ="Label12"
                            Caption ="EST Delv."
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =585
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6450
                    Top =285
                    Width =1125
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =4754549
                    ForeColor =4754549
                    Name ="UnitFinalCost"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6450
                    LayoutCachedTop =285
                    LayoutCachedWidth =7575
                    LayoutCachedHeight =540
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5865
                            Top =285
                            Width =525
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label13"
                            Caption ="Cost:"
                            BottomPadding =38
                            LayoutCachedLeft =5865
                            LayoutCachedTop =285
                            LayoutCachedWidth =6390
                            LayoutCachedHeight =540
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =1155
                    Width =945
                    Height =225
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="InServDate"
                    ControlSource ="UnitInServ"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =1155
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1380
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =105
                            Top =1155
                            Width =990
                            Height =225
                            FontSize =9
                            Name ="Label11"
                            Caption ="InServ:"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =1155
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1380
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3195
                    Top =870
                    Width =1035
                    Height =225
                    TabIndex =7
                    Name ="SchFSL"
                    ControlSource ="UnitTaxExempt"
                    StatusBarText ="FSL Lease"
                    GroupTable =6

                    LayoutCachedLeft =3195
                    LayoutCachedTop =870
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2160
                            Top =870
                            Width =975
                            Height =225
                            FontSize =9
                            Name ="Label6"
                            Caption ="TaxExempt:"
                            GroupTable =6
                            LayoutCachedLeft =2160
                            LayoutCachedTop =870
                            LayoutCachedWidth =3135
                            LayoutCachedHeight =1095
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3195
                    Top =585
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="VendorREF"
                    ControlSource ="VendorREF"
                    GroupTable =6

                    LayoutCachedLeft =3195
                    LayoutCachedTop =585
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =2160
                            Top =585
                            Width =975
                            Height =225
                            FontSize =8
                            Name ="Label17"
                            Caption ="Vendor"
                            GroupTable =6
                            LayoutCachedLeft =2160
                            LayoutCachedTop =585
                            LayoutCachedWidth =3135
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6450
                    Top =570
                    Width =1110
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =4754549
                    Name ="LastEstDelDate"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    RightPadding =38

                    LayoutCachedLeft =6450
                    LayoutCachedTop =570
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =825
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5865
                            Top =570
                            Width =525
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label14"
                            Caption ="Rent:"
                            LayoutCachedLeft =5865
                            LayoutCachedTop =570
                            LayoutCachedWidth =6390
                            LayoutCachedHeight =825
                            RowStart =2
                            RowEnd =2
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =870
                    Width =945
                    Height =225
                    FontSize =9
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483615
                    Name ="ActualDelDate"
                    ControlSource ="UnitActualDelDate"
                    StatusBarText ="deal id"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =870
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =105
                            Top =870
                            Width =990
                            Height =225
                            FontSize =9
                            Name ="Label10"
                            Caption ="Act Delv"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =870
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1095
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3195
                    Top =1155
                    Width =1035
                    Height =225
                    TabIndex =9
                    Name ="UnitInServInd"
                    ControlSource ="UnitInServInd"
                    GroupTable =6

                    LayoutCachedLeft =3195
                    LayoutCachedTop =1155
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =1380
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2160
                            Top =1155
                            Width =975
                            Height =225
                            FontSize =9
                            Name ="Label7"
                            Caption ="In Serv Ind"
                            GroupTable =6
                            LayoutCachedLeft =2160
                            LayoutCachedTop =1155
                            LayoutCachedWidth =3135
                            LayoutCachedHeight =1380
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =900
                    Width =525
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =16
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =0
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =900
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =1395
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4890
                    Top =1425
                    Width =345
                    Height =165
                    FontSize =8
                    TabIndex =17
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =4890
                    LayoutCachedTop =1425
                    LayoutCachedWidth =5235
                    LayoutCachedHeight =1590
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =300
                    Width =945
                    Height =225
                    FontSize =9
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483617
                    Name ="Text23"
                    ControlSource ="Unitbuilddate"
                    StatusBarText ="schedule number starts with 1"
                    ControlTipText ="Double click to open schedule"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =300
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =525
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =105
                            Top =300
                            Width =990
                            Height =225
                            FontSize =9
                            Name ="Label24"
                            Caption ="Build"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =300
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =525
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =2010
                    Width =945
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="Text27"
                    ControlSource ="unitoffleasedt"
                    Format ="Short Date"
                    GroupTable =6

                    LayoutCachedLeft =1155
                    LayoutCachedTop =2010
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2250
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =105
                            Top =2010
                            Width =990
                            Height =240
                            FontSize =8
                            Name ="Label28"
                            Caption ="Off Lease"
                            GroupTable =6
                            LayoutCachedLeft =105
                            LayoutCachedTop =2010
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =2250
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3195
                    Top =2010
                    Width =1035
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="VendorName"
                    GroupTable =6

                    LayoutCachedLeft =3195
                    LayoutCachedTop =2010
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =2250
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5145
                    Width =2479
                    FontSize =10
                    TabIndex =1
                    BackColor =15527148
                    BorderColor =0
                    Name ="Text55"
                    ControlSource ="UnitVIN"

                    LayoutCachedLeft =5145
                    LayoutCachedWidth =7624
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =0
                End
                Begin EmptyCell
                    Left =3195
                    Top =1440
                    Width =1035
                    Height =225
                    Name ="EmptyCell122"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =3195
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =1665
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =3195
                    Top =1725
                    Width =1035
                    Height =225
                    Name ="EmptyCell123"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =3195
                    LayoutCachedTop =1725
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2160
                    Top =1440
                    Width =975
                    Height =225
                    Name ="EmptyCell128"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3135
                    LayoutCachedHeight =1665
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2160
                    Top =1725
                    Width =975
                    Height =225
                    Name ="EmptyCell129"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =1725
                    LayoutCachedWidth =3135
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2160
                    Top =300
                    Width =975
                    Height =225
                    Name ="EmptyCell130"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =300
                    LayoutCachedWidth =3135
                    LayoutCachedHeight =525
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =3195
                    Top =300
                    Width =1035
                    Height =225
                    Name ="EmptyCell131"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =3195
                    LayoutCachedTop =300
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =525
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2160
                    Top =2010
                    Width =975
                    Name ="EmptyCell146"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =2010
                    LayoutCachedWidth =3135
                    LayoutCachedHeight =2250
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =6
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "UnitSnapShot.cls"
