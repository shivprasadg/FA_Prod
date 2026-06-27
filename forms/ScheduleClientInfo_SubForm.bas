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
    Width =4995
    DatasheetFontHeight =11
    ItemSuffix =160
    Left =10770
    Top =5130
    Right =15840
    Bottom =6630
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x121e1ff0903fe540
    End
    RecordSource ="Select * From ScheduleClientInfo Where ClientId=5295;"
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
            Height =1500
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =750
                    Width =4995
                    FontSize =8
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="AnnualSales"
                    ControlSource ="=[clientIndustry] & \"  - Annual Sales: \" & Format([clientAnnualSales],\"$#,###"
                        ".\")"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =750
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =990
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
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
                    Left =660
                    Top =990
                    Width =4335
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="clientMainPhone"
                    ControlSource ="clientMainPhone"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    InputMask ="!\\(999\") \"000\\-0000;0;_"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =660
                    LayoutCachedTop =990
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =1230
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
                            Top =990
                            Width =660
                            Height =240
                            FontSize =8
                            Name ="Label9"
                            Caption ="Phone "
                            GroupTable =6
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =990
                            LayoutCachedWidth =660
                            LayoutCachedHeight =1230
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =1230
                    Width =4335
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =11830108
                    Name ="clientWebsite"
                    ControlSource ="clientWebsite"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =660
                    LayoutCachedTop =1230
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =1470
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
                            Top =1230
                            Width =660
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Website "
                            GroupTable =6
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =1230
                            LayoutCachedWidth =660
                            LayoutCachedHeight =1470
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =540
                    Width =4995
                    Height =210
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="CityStateLine"
                    ControlSource ="=[clientCity] & \", \" & [clientState]"
                    StatusBarText ="deal id"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =540
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =750
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =1440
                    Width =1620
                    Height =60
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="clientID"
                    ControlSource ="clientID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =270
                    Width =4335
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =8210719
                    Name ="MasterLease"
                    StatusBarText ="schedule number starts with 1"
                    FontName ="Segoe UI Semibold"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double click to open schedule"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =660
                    LayoutCachedTop =270
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =540
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
                            Top =270
                            Width =660
                            Height =270
                            FontSize =10
                            FontWeight =700
                            ForeColor =6108695
                            Name ="Label146"
                            Caption ="Lease"
                            FontName ="Segoe UI Semibold"
                            GroupTable =6
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =270
                            LayoutCachedWidth =660
                            LayoutCachedHeight =540
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
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Width =4335
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =8210719
                    Name ="CompanyName"
                    ControlSource ="clientCompanyName"
                    StatusBarText ="schedule number starts with 1"
                    FontName ="Segoe UI Semibold"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double click to open schedule"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =660
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =270
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
                            Width =660
                            Height =270
                            FontSize =10
                            FontWeight =700
                            ForeColor =6108695
                            Name ="Label156"
                            Caption ="Client"
                            FontName ="Segoe UI Semibold"
                            GroupTable =6
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedWidth =660
                            LayoutCachedHeight =270
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
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
                    Left =1500
                    Top =1440
                    Width =1620
                    Height =60
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="MLID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
    End
End
CodeBehindForm
' See "ScheduleClientInfo_SubForm.cls"
