Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3631
    DatasheetFontHeight =11
    ItemSuffix =358
    Left =15030
    Top =10620
    Right =18450
    Bottom =14865
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x12dea8491604e540
    End
    RecordSource ="SELECT * From Units WHERE UnitID=323948;"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Section
            CanGrow = NotDefault
            Height =3915
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =2851
                    Width =945
                    FontSize =8
                    TabIndex =13
                    BackColor =-2147483643
                    BorderColor =32768
                    Name ="Text346"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =2851
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =3091
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =2851
                            Width =720
                            Height =240
                            FontSize =8
                            BackColor =32768
                            Name ="Label347"
                            Caption ="SyndRent"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =2851
                            LayoutCachedWidth =780
                            LayoutCachedHeight =3091
                            RowStart =2
                            RowEnd =2
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =2040
                    Width =945
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    Name ="Text27"
                    ControlSource ="unitoffleasedt"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =2040
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label28"
                            Caption ="Off Lease"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =2040
                            LayoutCachedWidth =780
                            LayoutCachedHeight =2280
                            RowStart =6
                            RowEnd =6
                            BackThemeColorIndex =0
                            BackTint =50.0
                            BorderThemeColorIndex =1
                            BorderShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =2310
                    Width =945
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    BorderColor =8965045
                    Name ="UnitFinalCost"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =2310
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =2550
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            Left =60
                            Top =2310
                            Width =720
                            Height =240
                            FontSize =8
                            BackColor =14151142
                            Name ="Label13"
                            Caption ="Cost"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =2310
                            LayoutCachedWidth =780
                            LayoutCachedHeight =2550
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Width =3598
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="UnitDesc"
                    ControlSource ="=[Unityr] & \" \" & [UnitMake] & \" \" & [UnitModel]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15
                    LayoutCachedWidth =3613
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =1290
                    Width =945
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="docSchedSent"
                    ControlSource ="UnitCOASent"
                    Format ="Short Date"
                    InputMask ="\"Sched Sent\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =1290
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =1530
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =1290
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label15"
                            Caption ="COA sent"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =1290
                            LayoutCachedWidth =780
                            LayoutCachedHeight =1530
                            RowStart =3
                            RowEnd =3
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =1530
                    Width =945
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    BorderColor =-2147483617
                    Name ="docSchedBack"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =1530
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =1770
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =60
                            Top =1530
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label16"
                            Caption ="Accepted"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =1530
                            LayoutCachedWidth =780
                            LayoutCachedHeight =1770
                            RowStart =4
                            RowEnd =4
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =810
                    Width =945
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="EstDelvDate"
                    ControlSource ="UnitEstDelDate"
                    Format ="Short Date"
                    StatusBarText ="schedule number starts with 1"
                    ControlTipText ="Double click to open schedule"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =810
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =1050
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =810
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label12"
                            Caption ="EST Delv."
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =810
                            LayoutCachedWidth =780
                            LayoutCachedHeight =1050
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =1770
                    Width =945
                    FontSize =8
                    TabIndex =9
                    BackColor =-2147483643
                    BorderColor =8872517
                    ForeColor =6108695
                    Name ="InServDate"
                    ControlSource ="UnitInServ"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =1770
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =2010
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =60
                            Top =1770
                            Width =720
                            Height =240
                            FontSize =8
                            BackColor =8872517
                            Name ="Label11"
                            Caption ="In Service"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =1770
                            LayoutCachedWidth =780
                            LayoutCachedHeight =2010
                            RowStart =5
                            RowEnd =5
                            BorderThemeColorIndex =2
                            BorderShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2640
                    Top =585
                    Width =30
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="VendorREF"
                    ControlSource ="VendorREF"

                    LayoutCachedLeft =2640
                    LayoutCachedTop =585
                    LayoutCachedWidth =2670
                    LayoutCachedHeight =795
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =1050
                    Width =945
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="ActualDelDate"
                    ControlSource ="UnitActualDelDate"
                    Format ="Short Date"
                    StatusBarText ="deal id"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =1050
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =60
                            Top =1050
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label10"
                            Caption ="Act Delv."
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =1050
                            LayoutCachedWidth =780
                            LayoutCachedHeight =1290
                            RowStart =2
                            RowEnd =2
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2430
                    Top =765
                    Width =555
                    Height =480
                    FontWeight =700
                    TabIndex =8
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BackColor =8872517
                    BorderColor =0
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2430
                    LayoutCachedTop =765
                    LayoutCachedWidth =2985
                    LayoutCachedHeight =1245
                    RowStart =1
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2925
                    Top =1125
                    Width =345
                    Height =165
                    FontSize =8
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =2925
                    LayoutCachedTop =1125
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =1290
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =570
                    Width =945
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="Text23"
                    ControlSource ="Unitbuilddate"
                    Format ="Short Date"
                    StatusBarText ="schedule number starts with 1"
                    ControlTipText ="Double click to open schedule"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =570
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =810
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =60
                            Top =570
                            Width =720
                            Height =240
                            FontSize =8
                            Name ="Label24"
                            Caption ="Build"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =570
                            LayoutCachedWidth =780
                            LayoutCachedHeight =810
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =795
                    Top =3120
                    Width =0
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =10040879
                    Name ="BuyerID"
                    ControlSource ="=IIf([UnitStatus] Like \"*R*\" Or [UnitStatus] Like \"*S*\",Trim(DLookUp(\"Buyer"
                        "Ref\",\"RMKT\",\"[UnitRef]=\" & [unitid] & \"\")),Null)"
                    Format ="General Number"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =795
                    LayoutCachedTop =3120
                    LayoutCachedWidth =795
                    LayoutCachedHeight =3330
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1890
                    Top =3615
                    Width =1620
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    ForeColor =0
                    Name ="CloneMe"
                    Caption ="Add Unit (Clone)"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Copy this Unit as a NEW Unit to the Group"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1890
                    LayoutCachedTop =3615
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =3855
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =90
                    Top =3615
                    Width =1620
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    ForeColor =0
                    Name ="AddBlankUnit"
                    Caption ="Add Unit (Blank)"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a NEW Unit to the Group"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =90
                    LayoutCachedTop =3615
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =3855
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2430
                    Top =585
                    Width =555
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="Label279"
                    Caption ="Status"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2430
                    LayoutCachedTop =585
                    LayoutCachedWidth =2985
                    LayoutCachedHeight =825
                    ColumnStart =2
                    ColumnEnd =3
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =765
                    Top =3091
                    Width =1200
                    Height =223
                    FontSize =8
                    TabIndex =15
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="ObcUnitNumber"
                    ControlSource ="ObcUnitNumber"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =765
                    LayoutCachedTop =3091
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =3314
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =45
                            Top =3091
                            Width =720
                            Height =223
                            FontSize =8
                            Name ="Label345"
                            Caption ="OBC#"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =45
                            LayoutCachedTop =3091
                            LayoutCachedWidth =765
                            LayoutCachedHeight =3314
                            RowStart =3
                            RowEnd =3
                            ForeThemeColorIndex =0
                            ForeTint =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =255
                    Width =3598
                    Height =255
                    FontSize =9
                    TabIndex =18
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="Text354"
                    ControlSource ="=[UnitSubType] & \" \" & [UnitType] & \" <Font Size=1>(<i>\" & DLookUp(\"VendorS"
                        "Nm\",\"Vendors\",\"VendorId=\" & [VendorREF] & \"\") & \"</i>)</font>\""
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =15
                    LayoutCachedTop =255
                    LayoutCachedWidth =3613
                    LayoutCachedHeight =510
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =2580
                    Width =945
                    FontSize =8
                    TabIndex =12
                    BackColor =-2147483643
                    BorderColor =32768
                    Name ="LastEstDelDate"
                    ControlSource ="UNITSRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =2580
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =2820
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =2580
                            Width =720
                            Height =240
                            FontSize =8
                            BackColor =8965045
                            Name ="Label14"
                            Caption ="Sch Rent"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =2580
                            LayoutCachedWidth =780
                            LayoutCachedHeight =2820
                            RowStart =2
                            RowEnd =2
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ScheduleUnitInfo_SubForm.cls"
