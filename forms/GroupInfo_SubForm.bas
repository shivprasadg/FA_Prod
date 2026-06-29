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
    Width =10426
    DatasheetFontHeight =11
    ItemSuffix =359
    Left =10770
    Top =6945
    Right =24705
    Bottom =9015
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x220f1cf0903fe540
    End
    RecordSource ="SELECT *, DLookUp('MLOrig','vw_SixKeys','SGrpID =' & SGrpID) As MLOrig FROM Schg"
        "rp WHERE Schgrp.SGrpID=2991;"
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =1815
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Top =750
                    Width =1065
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =6108695
                    Name ="LeasePymtForGroup"
                    ControlSource ="[LeasePymtForGroup]"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =750
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =990
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =750
                            Width =480
                            Height =240
                            FontSize =8
                            Name ="Label254"
                            Caption ="Paymt"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =750
                            LayoutCachedWidth =540
                            LayoutCachedHeight =990
                            RowStart =5
                            RowEnd =5
                            BackThemeColorIndex =1
                            BackShade =65.0
                            BorderThemeColorIndex =0
                            BorderTint =50.0
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
                    Left =540
                    Top =480
                    Width =1065
                    FontSize =8
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =6108695
                    Name ="OrigEquipCost"
                    ControlSource ="[OrigEquipCost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =480
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =720
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =480
                            Width =480
                            Height =240
                            FontSize =8
                            Name ="Label11"
                            Caption =" OEC"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =480
                            LayoutCachedWidth =540
                            LayoutCachedHeight =720
                            RowStart =5
                            RowEnd =5
                            BackThemeColorIndex =1
                            BackShade =50.0
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            ForeThemeColorIndex =1
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
                    Top =240
                    Width =4302
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="dealShortDesc"
                    ControlSource ="SgrpDesc"

                    LayoutCachedTop =240
                    LayoutCachedWidth =4302
                    LayoutCachedHeight =465
                    ThemeFontIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =4302
                    Height =270
                    FontSize =10
                    BorderColor =0
                    Name ="GroupNum"
                    ControlSource ="=\"Group \" & [SGrpNO]"
                    StatusBarText ="schedule number starts with 1"
                    FontName ="Segoe UI Semibold"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double click to open schedule"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =4302
                    LayoutCachedHeight =270
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3780
                    Top =645
                    Width =495
                    Height =345
                    FontWeight =700
                    TabIndex =3
                    BorderColor =16777215
                    Name ="AssignedIndicator\015\012\015\012\015\012\015\012"
                    ControlSource ="=IIf([mlorig]<>\"FA\",\"n/a\",[Assigned_Indicator])"
                    EventProcPrefix ="AssignedIndicator________"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =645
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =990
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =345
                    Height =165
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="SchID"
                    ControlSource ="SGrpSchID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =1200
                    Width =480
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Text244"
                    ControlSource ="unitsingroup"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1200
                    LayoutCachedWidth =540
                    LayoutCachedHeight =1500
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =60
                    Top =1035
                    Width =480
                    Height =240
                    FontSize =7
                    Name ="Label245"
                    Caption ="Units"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =1035
                    LayoutCachedWidth =540
                    LayoutCachedHeight =1275
                    BackThemeColorIndex =0
                    BackTint =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2520
                    Top =60
                    Width =345
                    Height =165
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="GroupID"
                    ControlSource ="SGrpID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =2520
                    LayoutCachedTop =60
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =1020
                    Width =3195
                    Height =225
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="Text248"
                    ControlSource ="=IIf([mlorig]<>\"FA\",Null,\"<B>\" & DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn"
                        "\",\"ID=\" & CInt([SgrpCommInd])) & \"</B> Set? <B>\" & IIf(IsGroupBLCDSet([grou"
                        "pid])=True,\"Y\",\"N\"))"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =1080
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =1245
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    Left =3780
                    Top =480
                    Width =495
                    Height =185
                    FontSize =7
                    BorderColor =16777215
                    Name ="Label66"
                    Caption ="Funded"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3780
                    LayoutCachedTop =480
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =665
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =1260
                    Width =3195
                    Height =225
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="Text251"
                    ControlSource ="=IIf([mlorig]<>\"FA\",Null,\"<B>\" & [PDStart]) & \"</B> Set? <B>\" & IIf(AllUni"
                        "tsInGroup_PerDiemSet([groupid])=True,\"Y\",\"N\")"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =1080
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =1485
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1200
                    Top =1500
                    Width =3075
                    Height =225
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="Text252"
                    ControlSource ="=\"All Units In Group Accepted? <B>\" & IIf(AllUnitsInGroup_Accepted([groupid])="
                        "True,\"Y\",\"N\")"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =1200
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =1725
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2610
                    Top =480
                    Width =854
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    BorderColor =4754549
                    ForeColor =6108695
                    Name ="Text255"
                    ControlSource ="[Assignee_Bank]"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2610
                    LayoutCachedTop =480
                    LayoutCachedWidth =3464
                    LayoutCachedHeight =720
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1950
                            Top =480
                            Width =660
                            Height =240
                            FontSize =8
                            BorderColor =4754549
                            Name ="Label256"
                            Caption ="Assignee"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =480
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =720
                            RowStart =5
                            RowEnd =5
                            BackThemeColorIndex =6
                            BackShade =50.0
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
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2610
                    Top =750
                    Width =854
                    FontSize =8
                    TabIndex =12
                    BackColor =-2147483643
                    BorderColor =4754549
                    ForeColor =6108695
                    Name ="Text257"
                    ControlSource ="[BLCD]"
                    Format ="mm/dd/yy"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2610
                    LayoutCachedTop =750
                    LayoutCachedWidth =3464
                    LayoutCachedHeight =990
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
                            OverlapFlags =93
                            TextAlign =1
                            Left =1950
                            Top =750
                            Width =660
                            Height =240
                            FontSize =8
                            BackColor =4754549
                            BorderColor =4754549
                            Name ="Label258"
                            Caption ="BLCD"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =750
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =990
                            RowStart =5
                            RowEnd =5
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =600
                    Top =1020
                    Width =480
                    Height =240
                    FontSize =8
                    BackColor =14211288
                    ForeColor =8210719
                    Name ="Label259"
                    Caption ="Basis"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =600
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1260
                    RowStart =5
                    RowEnd =5
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =600
                    Top =1260
                    Width =480
                    Height =240
                    FontSize =8
                    BackColor =14211288
                    ForeColor =8210719
                    Name ="Label260"
                    Caption ="PD on"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =600
                    LayoutCachedTop =1260
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1500
                    RowStart =5
                    RowEnd =5
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4845
                    Top =150
                    Width =5581
                    Height =1635
                    FontSize =8
                    TabIndex =13
                    Name ="TabCtl261"
                    FontName ="Calibri"

                    LayoutCachedLeft =4845
                    LayoutCachedTop =150
                    LayoutCachedWidth =10426
                    LayoutCachedHeight =1785
                    ThemeFontIndex =1
                    OldBorderStyle =0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =4921
                            Top =555
                            Width =5423
                            Height =1155
                            Name ="Economics"
                            LayoutCachedLeft =4921
                            LayoutCachedTop =555
                            LayoutCachedWidth =10344
                            LayoutCachedHeight =1710
                            Begin
                                Begin EmptyCell
                                    Left =4968
                                    Top =696
                                    Width =456
                                    Name ="EmptyCell354"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =4968
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =5424
                                    LayoutCachedHeight =936
                                    LayoutGroup =1
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =5436
                                    Top =696
                                    Width =612
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label619"
                                    Caption ="Index"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =6048
                                    LayoutCachedHeight =936
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =6060
                                    Top =696
                                    Width =768
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label282"
                                    Caption =" Date"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =6828
                                    LayoutCachedHeight =936
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =6840
                                    Top =696
                                    Width =576
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label284"
                                    Caption ="Rate"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =6840
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =7416
                                    LayoutCachedHeight =936
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =2
                                    Left =7428
                                    Top =696
                                    Width =420
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label283"
                                    Caption ="Year"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =7428
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =7848
                                    LayoutCachedHeight =936
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =2
                                    Left =7848
                                    Top =696
                                    Width =588
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label286"
                                    Caption ="Spread"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =7848
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =8436
                                    LayoutCachedHeight =936
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =2
                                    Left =8448
                                    Top =696
                                    Width =516
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label285"
                                    Caption ="Yield"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =8448
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =8964
                                    LayoutCachedHeight =936
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =8976
                                    Top =696
                                    Width =612
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label299"
                                    Caption ="Bonus"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =8976
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =936
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =9600
                                    Top =696
                                    Width =588
                                    Height =240
                                    FontSize =8
                                    BorderColor =16777215
                                    Name ="Label347"
                                    Caption ="Adjmt"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =9600
                                    LayoutCachedTop =696
                                    LayoutCachedWidth =10188
                                    LayoutCachedHeight =936
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6060
                                    Top =948
                                    Width =768
                                    FontSize =8
                                    Name ="Text325"
                                    ControlSource ="SGrpSalesIndDt"
                                    Format ="mm/dd/yy"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =6828
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6840
                                    Top =948
                                    Width =576
                                    FontSize =8
                                    Name ="Text340"
                                    ControlSource ="SgrpSalesIndRt"
                                    Format ="Percent"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =6840
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =7416
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    OverlapFlags =223
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7428
                                    Top =948
                                    Width =420
                                    FontSize =8
                                    Name ="Text342"
                                    ControlSource ="SGprSalesIndYr"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =7428
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =7848
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    DecimalPlaces =2
                                    SpecialEffect =5
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7848
                                    Top =948
                                    Width =588
                                    FontSize =8
                                    Name ="Text344"
                                    ControlSource ="=[SgrpSalesYield]-[SgrpSalesIndRt]"
                                    Format ="Percent"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                                        0x00000000ffc20e00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =7848
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =8436
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =85.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000050000000100000000000000ffc20e00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =255
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8448
                                    Top =948
                                    Width =516
                                    FontSize =8
                                    Name ="Text346"
                                    ControlSource ="SgrpSalesYield"
                                    Format ="Percent"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =8448
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =8964
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9600
                                    Top =948
                                    Width =588
                                    FontSize =8
                                    Name ="Text350"
                                    ControlSource ="SgrpYield"
                                    Format ="Percent"
                                    StatusBarText ="downadj"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Rate Adjustment - If Downard"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =9600
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =10188
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =6060
                                    Top =1200
                                    Width =768
                                    FontSize =8
                                    Name ="SwapDate"
                                    ControlSource ="SGrpIndDt"
                                    Format ="mm/dd/yy"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =6828
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =6840
                                    Top =1200
                                    Width =576
                                    FontSize =8
                                    Name ="SwapRate"
                                    ControlSource ="IndexRate"
                                    Format ="Percent"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =6840
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =7416
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7428
                                    Top =1200
                                    Width =420
                                    FontSize =8
                                    Name ="SwapTermYears"
                                    ControlSource ="SGprIndYr"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =7428
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =7848
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    SpecialEffect =5
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7848
                                    Top =1200
                                    Width =588
                                    FontSize =8
                                    Name ="SalesSpread"
                                    ControlSource ="=[sgrpyield]-[IndexRate]"
                                    Format ="Percent"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                                        0x00000000ffc20e00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =7848
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =8436
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =85.0
                                    ThemeFontIndex =1
                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000050000000100000000000000ffc20e00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =8448
                                    Top =1200
                                    Width =516
                                    FontSize =8
                                    Name ="SwapYield"
                                    ControlSource ="SgrpYield"
                                    Format ="Percent"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =8448
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =8964
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    DecimalPlaces =3
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =9600
                                    Top =1200
                                    Width =588
                                    FontSize =8
                                    Name ="SwapAdjustment"
                                    ControlSource ="SGrpBonus"
                                    Format ="Percent"
                                    StatusBarText ="downadj"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Rate Adjustment - If Downard"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =9600
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =10188
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =1200
                                    Width =612
                                    FontSize =8
                                    Name ="SwapIndex"
                                    ControlSource ="SGrpIndex"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =6048
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextAlign =2
                                            Left =4968
                                            Top =1200
                                            Width =456
                                            Height =240
                                            FontSize =8
                                            BackColor =8965045
                                            BorderColor =16777215
                                            Name ="Label357"
                                            Caption ="Synd"
                                            GroupTable =2
                                            LeftPadding =0
                                            TopPadding =0
                                            RightPadding =0
                                            BottomPadding =0
                                            LayoutCachedLeft =4968
                                            LayoutCachedTop =1200
                                            LayoutCachedWidth =5424
                                            LayoutCachedHeight =1440
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =1
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =2
                                        End
                                    End
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =948
                                    Width =612
                                    FontSize =8
                                    Name ="Combo324"
                                    ControlSource ="SGrpSalesIndex"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =6048
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextAlign =2
                                            Left =4968
                                            Top =948
                                            Width =456
                                            Height =240
                                            FontSize =8
                                            BorderColor =16777215
                                            Name ="Label352"
                                            Caption ="LO"
                                            GroupTable =2
                                            LeftPadding =0
                                            TopPadding =0
                                            RightPadding =0
                                            BottomPadding =0
                                            LayoutCachedLeft =4968
                                            LayoutCachedTop =948
                                            LayoutCachedWidth =5424
                                            LayoutCachedHeight =1188
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =1
                                            ThemeFontIndex =1
                                            BackThemeColorIndex =1
                                            BackShade =85.0
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =2
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =8976
                                    Top =1200
                                    Width =612
                                    FontSize =8
                                    Name ="SwapBouns"
                                    ControlSource ="SGrpBonus"
                                    StatusBarText ="Bonus Depreciation"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =8976
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =1440
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                                Begin TextBox
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8976
                                    Top =948
                                    Width =612
                                    FontSize =8
                                    Name ="Combo348"
                                    ControlSource ="SGrpSalesBonus"
                                    StatusBarText ="Bonus Depreciation"
                                    GroupTable =2
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =8976
                                    LayoutCachedTop =948
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =1188
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =75.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =2
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =4921
                            Top =555
                            Width =5430
                            Height =1155
                            Name ="Future Use"
                            EventProcPrefix ="Future_Use"
                            LayoutCachedLeft =4921
                            LayoutCachedTop =555
                            LayoutCachedWidth =10351
                            LayoutCachedHeight =1710
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GroupInfo_SubForm.cls"
