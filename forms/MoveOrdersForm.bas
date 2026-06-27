Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =1
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5400
    DatasheetFontHeight =11
    ItemSuffix =402
    Left =4845
    Top =3315
    Right =12030
    Bottom =13440
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xb90749465431e640
    End
    RecordSource ="UnitMoveOrder"
    OnClose ="[Event Procedure]"
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
        Begin Image
            BackStyle =0
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
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
            Height =600
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =360
                    Width =5400
                    Height =210
                    BorderColor =16777215
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =360
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =570
                    TabIndex =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =120
                    Width =1785
                    Height =495
                    FontSize =16
                    Name ="Label1053"
                    Caption ="Move Order"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =495
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="mAssetId"
                    ControlSource ="mAssetId"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =1920
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =240
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =3780
                    Top =60
                    Width =1560
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="AddNew"
                    Caption ="Add a Move Order"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
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
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =180
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="mFaUserID"
                    ControlSource ="mFaUserID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =1920
                    LayoutCachedTop =180
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =180
                    Width =1020
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="UnitMoveID"
                    ControlSource ="UnitMoveID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =2700
                    LayoutCachedTop =180
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Visible = NotDefault
            CanGrow = NotDefault
            Height =6960
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =1260
                    Top =2160
                    Width =3960
                    Height =300
                    FontWeight =700
                    Name ="mOldLocationID"
                    ControlSource ="mOldLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                        "ete)=No) AND ((Locations.ClientGroupID)=[FORMS]![UnitsForm]![CLIENTGROUPID])) OR"
                        " (((Locations.ClientGroupID)=771)) ORDER BY Locations.LocationName, Locations.Lo"
                        "cationType;"
                    ColumnWidths ="0;2160;2160;1440;720;720"
                    StatusBarText ="Link to Contacts table"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2460
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =7
                    BackTint =40.0
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =240
                            Top =2160
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label5170"
                            Caption ="Move From"
                            BottomPadding =150
                            LayoutCachedLeft =240
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =2460
                            ThemeFontIndex =1
                            BackThemeColorIndex =7
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =1260
                    Top =3600
                    Width =3960
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="mNewLocationID"
                    ControlSource ="mNewLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                        "ete)=No) AND ((Locations.ClientGroupID)=[FORMS]![UnitsForm]![CLIENTGROUPID])) OR"
                        " (((Locations.ClientGroupID)=771)) ORDER BY Locations.LocationName, Locations.Lo"
                        "cationType;"
                    ColumnWidths ="0;2160;2160;1440;720;720"
                    StatusBarText ="Link to Contacts table"
                    AfterUpdate ="[Event Procedure]"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3600
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3900
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =240
                            Top =3600
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label5165"
                            Caption ="Move To"
                            BottomPadding =150
                            LayoutCachedLeft =240
                            LayoutCachedTop =3600
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3900
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =6
                            BackShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1260
                    Top =1740
                    Width =3960
                    Height =300
                    TabIndex =2
                    BackColor =16183539
                    Name ="iVendorID"
                    ControlSource ="mVendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorOb"
                        "solete FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.Vendo"
                        "rSNm;"
                    ColumnWidths ="0;10800;2880"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1740
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2040
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =540
                            Top =1740
                            Width =660
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label76"
                            Caption ="Vendor"
                            BottomPadding =150
                            LayoutCachedLeft =540
                            LayoutCachedTop =1740
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =2040
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =2520
                    Width =3960
                    Height =300
                    TabIndex =3
                    Name ="mContact"
                    ControlSource ="mContact"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2820
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =300
                            Top =2520
                            Width =900
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label375"
                            Caption ="Contact "
                            BottomPadding =150
                            LayoutCachedLeft =300
                            LayoutCachedTop =2520
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =2820
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1260
                    Top =5460
                    Height =300
                    TabIndex =5
                    Name ="wStartDate"
                    ControlSource ="mStartDate"
                    Format ="Short Date"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"wStartDate\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">AddRmktExpFor"
                        End
                        Begin
                            Comment ="_AXL:m</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =1260
                    LayoutCachedTop =5460
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =5760
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =240
                            Top =5460
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label108"
                            Caption ="Start Date"
                            LayoutCachedLeft =240
                            LayoutCachedTop =5460
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =5760
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =4
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =1380
                    Width =3960
                    Height =300
                    TabIndex =4
                    Name ="mOrderNum"
                    ControlSource ="mOrderNum"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1680
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =1380
                            Width =1080
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Move Order "
                            LayoutCachedLeft =120
                            LayoutCachedTop =1380
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =1680
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3780
                    Top =5460
                    Height =300
                    TabIndex =6
                    Name ="mEndDate"
                    ControlSource ="mEndDate"
                    Format ="Short Date"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"mEndDate\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">AddRmktExpForm<"
                        End
                        Begin
                            Comment ="_AXL:/Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =5460
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =5760
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2820
                            Top =5460
                            Width =900
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label285"
                            Caption ="Close Date"
                            LayoutCachedLeft =2820
                            LayoutCachedTop =5460
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =5760
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =6
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =3960
                    Width =3960
                    Height =1020
                    TabIndex =7
                    BackColor =13434879
                    Name ="mNote"
                    ControlSource ="mNote"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =4980
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =600
                            Top =3960
                            Width =600
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label152"
                            Caption ="Notes "
                            LayoutCachedLeft =600
                            LayoutCachedTop =3960
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =4260
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =2880
                    Width =3960
                    Height =300
                    TabIndex =8
                    Name ="mEmail"
                    ControlSource ="mEmail"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =2880
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3180
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =540
                            Top =2880
                            Width =660
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label394"
                            Caption ="Email "
                            BottomPadding =150
                            LayoutCachedLeft =540
                            LayoutCachedTop =2880
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3180
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =3240
                    Width =1980
                    Height =300
                    TabIndex =10
                    Name ="mPhone"
                    ControlSource ="mPhone"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3240
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3540
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =540
                            Top =3240
                            Width =720
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label396"
                            Caption ="Phone "
                            BottomPadding =150
                            LayoutCachedLeft =540
                            LayoutCachedTop =3240
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =3540
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3780
                    Top =5880
                    Height =300
                    TabIndex =11
                    Name ="mMoveCost"
                    ControlSource ="mMoveCost"
                    BottomPadding =150

                    LayoutCachedLeft =3780
                    LayoutCachedTop =5880
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =6180
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =2820
                            Top =5880
                            Width =900
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label400"
                            Caption ="Cost "
                            BottomPadding =150
                            LayoutCachedLeft =2820
                            LayoutCachedTop =5880
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =6180
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =2
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3420
                    Top =5100
                    Width =1800
                    Height =240
                    FontSize =9
                    TabIndex =12
                    ForeColor =0
                    Name ="SubmitMoveOrder"
                    Caption ="Submit Order"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3420
                    LayoutCachedTop =5100
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =5340
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =5040
                    Width =1800
                    Height =315
                    TabIndex =9
                    Name ="mStatus"
                    ControlSource ="mStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"Not Ordred\";\"Pending\";\"In Transit\";\"Done\""
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =5040
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =5355
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =240
                            Top =5040
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label398"
                            Caption ="Status "
                            BottomPadding =150
                            LayoutCachedLeft =240
                            LayoutCachedTop =5040
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =5340
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =4
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =240
                    Top =60
                    Width =4974
                    Height =1260
                    FontSize =9
                    TabIndex =13
                    Name ="MoveOrderList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMoveID , mStartDate AS [Start], mEndDate AS [End], PickupLocation As "
                        "PickUp, DeliveryLocation AS Deliver, MStatus, AssetID FROM vw_UnitMoveDashboard "
                        "WHERE AssetID = 363218 ORDER BY mStartDate;"
                    ColumnWidths ="0;1008;1008;1440;1440;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =5214
                    LayoutCachedHeight =1320
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =240
                    Top =6300
                    Width =2460
                    Height =240
                    FontSize =9
                    TabIndex =14
                    Name ="DeleteLocation"
                    Caption ="Delete  Selected Move"
                    OnClick ="[Event Procedure]"
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =240
                    LayoutCachedTop =6300
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =6540
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =5
                    BorderTint =100.0
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
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
' See "MoveOrdersForm.cls"
