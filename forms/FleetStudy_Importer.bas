Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =23100
    DatasheetFontHeight =10
    ItemSuffix =3098
    Left =8985
    Top =5280
    Right =32340
    Bottom =20415
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x3c4800759370e540
    End
    RecordSource ="Select * From Clients Where ClientID=0"
    Caption ="Import Fleet Study Units"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Company Navigation"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =990
            Name ="FormHeader"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5520
                    Top =165
                    Width =9255
                    Height =465
                    ColumnOrder =5
                    FontSize =18
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8210719
                    Name ="ClientName"
                    ControlSource ="ClientShNm"
                    FontName ="Century Gothic"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =165
                    LayoutCachedWidth =14775
                    LayoutCachedHeight =630
                End
                Begin Label
                    OverlapFlags =247
                    Left =5520
                    Width =2040
                    Height =240
                    FontSize =9
                    Name ="Label683"
                    Caption ="You're working with:"
                    FontName ="Calibri"
                    LayoutCachedLeft =5520
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =240
                    ThemeFontIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19035
                    Top =60
                    Width =780
                    Height =210
                    ColumnOrder =4
                    FontSize =9
                    Name ="clientID"
                    ControlSource ="clientID"
                    FontName ="Calibri"
                    GroupTable =50
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19035
                    LayoutCachedTop =60
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =270
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =50
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =18300
                            Top =60
                            Width =735
                            Height =210
                            Name ="Label641"
                            Caption ="ClientID:"
                            FontName ="Calibri"
                            GroupTable =50
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18300
                            LayoutCachedTop =60
                            LayoutCachedWidth =19035
                            LayoutCachedHeight =270
                            LayoutGroup =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =50
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Top =690
                    Width =1710
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label228"
                    Caption ="C O R E"
                    FontName ="Segoe UI"
                    GroupTable =46
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =690
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =930
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =46
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1710
                    Width =2850
                    Height =690
                    FontSize =24
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Fleet Import"
                    FontName ="Segoe UI"
                    GroupTable =46
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1710
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =46
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Width =1710
                    Height =690
                    FontSize =24
                    Name ="Label466"
                    Caption ="ATLAAS"
                    FontName ="Segoe UI"
                    GroupTable =46
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =690
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =46
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1710
                    Top =690
                    Width =2850
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label227"
                    Caption ="DASH BOARD"
                    FontName ="Segoe UI"
                    GroupTable =46
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1710
                    LayoutCachedTop =690
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =930
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =46
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19035
                    Top =510
                    Width =780
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =2
                    Name ="SchID"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"
                    GroupTable =50
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19035
                    LayoutCachedTop =510
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =750
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =50
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =18300
                            Top =510
                            Width =735
                            Height =240
                            Name ="Label2877"
                            Caption ="SchID:"
                            FontName ="Calibri"
                            GroupTable =50
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18300
                            LayoutCachedTop =510
                            LayoutCachedWidth =19035
                            LayoutCachedHeight =750
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =50
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19035
                    Top =270
                    Width =780
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =1
                    Name ="MLID"
                    FontName ="Calibri"
                    GroupTable =50
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19035
                    LayoutCachedTop =270
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =510
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =50
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =18300
                            Top =270
                            Width =735
                            Height =240
                            Name ="Label1939"
                            Caption ="MLID:"
                            FontName ="Calibri"
                            GroupTable =50
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18300
                            LayoutCachedTop =270
                            LayoutCachedWidth =19035
                            LayoutCachedHeight =510
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =50
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19035
                    Top =750
                    Width =780
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =3
                    Name ="GroupID"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"
                    GroupTable =50
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19035
                    LayoutCachedTop =750
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =990
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =50
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =18300
                            Top =750
                            Width =735
                            Height =240
                            Name ="Label2710"
                            Caption =" GroupID:"
                            FontName ="Calibri"
                            GroupTable =50
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18300
                            LayoutCachedTop =750
                            LayoutCachedWidth =19035
                            LayoutCachedHeight =990
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =50
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =21795
                    Top =120
                    Width =750
                    Height =228
                    TabIndex =5
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =21795
                    LayoutCachedTop =120
                    LayoutCachedWidth =22545
                    LayoutCachedHeight =348
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
                    Left =21240
                    Top =120
                    Width =555
                    Height =228
                    TabIndex =6
                    Name ="GoHome"
                    Caption ="Home"
                    FontName ="Calibri"
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

                    LayoutCachedLeft =21240
                    LayoutCachedTop =120
                    LayoutCachedWidth =21795
                    LayoutCachedHeight =348
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =14160
            Name ="Detail"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =2895
                    Top =5880
                    Width =2385
                    Height =3900
                    BackColor =15527148
                    Name ="Box570"
                    LayoutCachedLeft =2895
                    LayoutCachedTop =5880
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =9780
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =2970
                    Top =5760
                    Width =1740
                    Height =240
                    FontSize =9
                    Name ="Label571"
                    Caption ="Import Functions"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2970
                    LayoutCachedTop =5760
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =6000
                    RowStart =2
                    RowEnd =2
                    BackThemeColorIndex =2
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2895
                    Top =75
                    Width =2475
                    Height =420
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label120"
                    Caption ="     Company Details"
                    FontName ="Calibri"
                    LayoutCachedLeft =2895
                    LayoutCachedTop =75
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =495
                    RowEnd =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =555
                    Width =3360
                    Height =270
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =967423
                    Name ="CompanyName"
                    ControlSource ="clientCompanyName"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."

                    LayoutCachedLeft =3795
                    LayoutCachedTop =555
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =825
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2895
                            Top =555
                            Width =840
                            Height =270
                            BorderColor =967423
                            Name ="Label138"
                            Caption ="Name:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2895
                            LayoutCachedTop =555
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =825
                            RowStart =2
                            RowEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =885
                    Width =3360
                    FontSize =9
                    TabIndex =9
                    BorderColor =11830108
                    Name ="clientAddress1"
                    ControlSource ="clientAddress1"
                    FontName ="Calibri"

                    LayoutCachedLeft =3795
                    LayoutCachedTop =885
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =1125
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2895
                            Top =885
                            Width =840
                            Height =240
                            Name ="Label140"
                            Caption ="Address1:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2895
                            LayoutCachedTop =885
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =1125
                            RowStart =3
                            RowEnd =3
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =1185
                    Width =3360
                    FontSize =9
                    TabIndex =12
                    BorderColor =11830108
                    Name ="clientAddress2"
                    ControlSource ="clientAddress2"
                    FontName ="Calibri"

                    LayoutCachedLeft =3795
                    LayoutCachedTop =1185
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =1425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2895
                            Top =1185
                            Width =840
                            Height =240
                            Name ="Label141"
                            Caption ="Address2:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2895
                            LayoutCachedTop =1185
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =1425
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =1485
                    Width =1575
                    FontSize =9
                    TabIndex =15
                    BorderColor =11830108
                    Name ="clientCity"
                    ControlSource ="clientCity"
                    FontName ="Calibri"

                    LayoutCachedLeft =3795
                    LayoutCachedTop =1485
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =1725
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2895
                            Top =1485
                            Width =840
                            Height =240
                            Name ="Label142"
                            Caption ="City/St/Zip:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2895
                            LayoutCachedTop =1485
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =1725
                            RowStart =5
                            RowEnd =5
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5430
                    Top =1485
                    Width =690
                    FontSize =9
                    TabIndex =16
                    BorderColor =11830108
                    Name ="clientState"
                    ControlSource ="clientState"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_States"
                    StatusBarText ="lookup states table"
                    FontName ="Calibri"

                    LayoutCachedLeft =5430
                    LayoutCachedTop =1485
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =1725
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6180
                    Top =1485
                    Width =975
                    FontSize =9
                    TabIndex =17
                    BorderColor =11830108
                    Name ="clientZip"
                    ControlSource ="clientZip"
                    FontName ="Calibri"
                    InputMask ="00000\\-9999;;_"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =1485
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =1725
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3555
                    Top =3870
                    Width =1650
                    FontSize =9
                    TabIndex =34
                    Name ="clientMainPhone"
                    ControlSource ="clientMainPhone"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    LeftPadding =60

                    LayoutCachedLeft =3555
                    LayoutCachedTop =3870
                    LayoutCachedWidth =5205
                    LayoutCachedHeight =4110
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =6
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2880
                            Top =3870
                            Width =585
                            Height =240
                            Name ="Label27"
                            Caption ="Main Phone:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =3870
                            LayoutCachedWidth =3465
                            LayoutCachedHeight =4110
                            RowStart =3
                            RowEnd =3
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3555
                    Top =4170
                    Width =3570
                    FontSize =9
                    TabIndex =38
                    Name ="clientWebsite"
                    ControlSource ="clientWebsite"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =3555
                    LayoutCachedTop =4170
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =4410
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =10
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2880
                            Top =4170
                            Width =585
                            Height =240
                            Name ="Label35"
                            Caption ="Website:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =4170
                            LayoutCachedWidth =3465
                            LayoutCachedHeight =4410
                            RowStart =4
                            RowEnd =4
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3810
                    Top =1860
                    Width =1575
                    FontSize =9
                    TabIndex =19
                    Name ="LeaseSTofCorp"
                    ControlSource ="LeaseSTofCorp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev, lookup_States.StateLong FROM lookup_States ORD"
                        "ER BY lookup_States.StateAbbrev;"
                    ColumnWidths ="0;1728"
                    FontName ="Calibri"
                    BottomPadding =120

                    LayoutCachedLeft =3810
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =2100
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3795
                    Top =2175
                    Width =1575
                    FontSize =9
                    TabIndex =26
                    Name ="clientTaxID"
                    ControlSource ="clientTaxID"
                    FontName ="Calibri"
                    BottomPadding =120

                    LayoutCachedLeft =3795
                    LayoutCachedTop =2175
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =2415
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2895
                            Top =2175
                            Width =840
                            Height =240
                            Name ="Label248"
                            Caption ="Tax ID:"
                            FontName ="Calibri"
                            BottomPadding =120
                            LayoutCachedLeft =2895
                            LayoutCachedTop =2175
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =2415
                            RowStart =7
                            RowEnd =7
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6180
                    Top =2175
                    Width =975
                    ColumnWidth =2535
                    FontSize =9
                    TabIndex =27
                    Name ="CLIENTGRPCNTRY"
                    ControlSource ="CLIENTGRPCNTRY"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Country.CNID, Country.Country FROM Country;"
                    ColumnWidths ="0;1440"
                    FontName ="Calibri"
                    BottomPadding =120
                    AllowValueListEdits =1

                    LayoutCachedLeft =6180
                    LayoutCachedTop =2175
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =2415
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3555
                    Top =3540
                    Width =1650
                    Height =270
                    FontSize =9
                    TabIndex =31
                    BackColor =13434879
                    BorderColor =967423
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    StatusBarText ="Client Short Nm Required (record will not be saved w/o)"
                    ValidationRule ="Is Not Null"
                    ValidationText ="A Short Name  is Required"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Client Short Nm Required (record will not be saved w/o)"
                    LeftPadding =60

                    LayoutCachedLeft =3555
                    LayoutCachedTop =3540
                    LayoutCachedWidth =5205
                    LayoutCachedHeight =3810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =6
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2880
                            Top =3540
                            Width =585
                            Height =270
                            BorderColor =967423
                            Name ="Label358"
                            Caption ="Short Name"
                            FontName ="Calibri"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =3540
                            LayoutCachedWidth =3465
                            LayoutCachedHeight =3810
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5955
                    Top =3540
                    Width =1170
                    Height =270
                    FontSize =9
                    TabIndex =32
                    BackColor =13434879
                    BorderColor =967423
                    Name ="clientType"
                    ControlSource ="clientType"
                    RowSourceType ="Value List"
                    RowSource ="TARGET;2ND TARGET;3RD TARGET;DEAD;CLIENT;GUARANTOR"
                    StatusBarText ="ClientType Required (record will not be saved w/o)"
                    ValidationRule ="Is Not Null"
                    ValidationText ="CLIENT TYPE REQUIRED"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="ClientType Required (record will not be saved w/o)"
                    RightPadding =75

                    LayoutCachedLeft =5955
                    LayoutCachedTop =3540
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =3810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =10
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1800
                    Left =5955
                    Top =3870
                    Width =1170
                    FontSize =9
                    TabIndex =35
                    BackColor =13434879
                    BorderColor =967423
                    Name ="LOinitial"
                    ControlSource ="LOinitial"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.OBSOLETE)=False) AND ((Em"
                        "ps.[LO Exec])=True)) ORDER BY Emps.EmpInt;"
                    ColumnWidths ="360;1440"
                    StatusBarText ="actual initials replaces reference"
                    ValidationRule ="Is Not Null"
                    ValidationText ="LO Exec Required"
                    DefaultValue ="\"HS\""
                    FontName ="Calibri"
                    Tag ="Required"
                    RightPadding =75

                    LayoutCachedLeft =5955
                    LayoutCachedTop =3870
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =4110
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =10
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =5265
                            Top =3870
                            Width =630
                            Height =240
                            BorderColor =967423
                            Name ="Label585"
                            Caption ="VP, BD:"
                            FontName ="Calibri"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =3870
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =4110
                            RowStart =3
                            RowEnd =3
                            ColumnStart =7
                            ColumnEnd =8
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =825
                    Width =2715
                    Height =12795
                    FontSize =8
                    TabIndex =8
                    BoundColumn =1
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Client, clientID FROM ClientsList WHERE Client Like '*xy*' ORDER BY Clien"
                        "t;"
                    ColumnWidths ="2016;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =825
                    LayoutCachedWidth =2790
                    LayoutCachedHeight =13620
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =450
                    Width =2715
                    Height =270
                    FontSize =9
                    TabIndex =2
                    BorderColor =967423
                    Name ="SearchClient"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =450
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =720
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =45
                            Top =75
                            Width =2895
                            Height =300
                            FontSize =9
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label668"
                            Caption ="Seach / Select Exisitng Companies"
                            FontName ="Calibri"
                            LayoutCachedLeft =45
                            LayoutCachedTop =75
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =375
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =5265
                    Top =3540
                    Width =630
                    Height =270
                    BorderColor =967423
                    Name ="Label37"
                    Caption ="Type:"
                    FontName ="Calibri"
                    LayoutCachedLeft =5265
                    LayoutCachedTop =3540
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =3810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =8
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5430
                    Top =2175
                    Width =690
                    Height =240
                    Name ="Label333"
                    Caption ="Foreign?:"
                    FontName ="Calibri"
                    BottomPadding =120
                    LayoutCachedLeft =5430
                    LayoutCachedTop =2175
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =2415
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5430
                    Top =75
                    Width =690
                    Height =228
                    BackColor =62207
                    BorderColor =62207
                    ForeColor =3381759
                    Name ="LockUnLock"
                    Caption ="UnLocked"
                    BottomPadding =120
                    LayoutCachedLeft =5430
                    LayoutCachedTop =75
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =303
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3555
                    Top =3060
                    Width =3570
                    Height =228
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    Name ="AddNewCompany"
                    Caption ="Add A New Company"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3555
                    LayoutCachedTop =3060
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =3288
                    ColumnStart =5
                    ColumnEnd =10
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8872517
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =7473
                    Top =603
                    Width =2490
                    Height =1470
                    FontSize =9
                    TabIndex =4
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig, Client, MLDt, MLType From MstrLease WHERE Client = 52"
                        "15 ORDER BY MstrLease.MLNo;"
                    ColumnWidths ="0;1728;576;0;0"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    AllowValueListEdits =0

                    LayoutCachedLeft =7473
                    LayoutCachedTop =603
                    LayoutCachedWidth =9963
                    LayoutCachedHeight =2073
                    RowStart =2
                    RowEnd =6
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7473
                    Top =123
                    Width =2490
                    Height =420
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label1860"
                    Caption ="     Select Master Lease"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    LayoutCachedLeft =7473
                    LayoutCachedTop =123
                    LayoutCachedWidth =9963
                    LayoutCachedHeight =543
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10833
                    Top =123
                    Width =2130
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="AddMasterLease"
                    Caption ="Add A Master Lease "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10833
                    LayoutCachedTop =123
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =351
                    ColumnStart =14
                    ColumnEnd =15
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8872517
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10833
                    Top =933
                    Width =2130
                    FontSize =9
                    TabIndex =10
                    BackColor =13952764
                    BorderColor =11830108
                    Name ="MLNo"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =10833
                    LayoutCachedTop =933
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =1173
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =15
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =10068
                            Top =933
                            Width =675
                            Height =240
                            Name ="Label1863"
                            Caption ="MLNo:"
                            FontName ="Calibri"
                            LayoutCachedLeft =10068
                            LayoutCachedTop =933
                            LayoutCachedWidth =10743
                            LayoutCachedHeight =1173
                            RowStart =3
                            RowEnd =3
                            ColumnStart =13
                            ColumnEnd =13
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10833
                    Top =603
                    Width =2130
                    Height =270
                    FontSize =9
                    TabIndex =5
                    BackColor =13952764
                    BorderColor =11830108
                    Name ="MLOrg"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =10833
                    LayoutCachedTop =603
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =873
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =15
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =10068
                            Top =603
                            Width =675
                            Height =270
                            Name ="Label1865"
                            Caption ="MLOrg:"
                            FontName ="Calibri"
                            LayoutCachedLeft =10068
                            LayoutCachedTop =603
                            LayoutCachedWidth =10743
                            LayoutCachedHeight =873
                            RowStart =2
                            RowEnd =2
                            ColumnStart =13
                            ColumnEnd =13
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10833
                    Top =1233
                    Width =2130
                    FontSize =9
                    TabIndex =13
                    BackColor =13952764
                    BorderColor =11830108
                    Name ="MLDt"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =10833
                    LayoutCachedTop =1233
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =1473
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =15
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =10068
                            Top =1233
                            Width =675
                            Height =240
                            Name ="Label1891"
                            Caption ="MLDate:"
                            FontName ="Calibri"
                            LayoutCachedLeft =10068
                            LayoutCachedTop =1233
                            LayoutCachedWidth =10743
                            LayoutCachedHeight =1473
                            RowStart =4
                            RowEnd =4
                            ColumnStart =13
                            ColumnEnd =13
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10833
                    Top =1833
                    Width =1072
                    Height =240
                    FontSize =9
                    TabIndex =21
                    Name ="SubmitMasterLease"
                    Caption ="Insert"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10833
                    LayoutCachedTop =1833
                    LayoutCachedWidth =11905
                    LayoutCachedHeight =2073
                    RowStart =6
                    RowEnd =6
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7473
                    Top =2073
                    Width =2490
                    Height =240
                    FontSize =9
                    TabIndex =25
                    Name ="DeleteMLease"
                    Caption ="Delete Selected Lease"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7473
                    LayoutCachedTop =2073
                    LayoutCachedWidth =9963
                    LayoutCachedHeight =2313
                    RowStart =9
                    RowEnd =9
                    ColumnStart =12
                    ColumnEnd =12
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17865
                    Top =933
                    Width =2415
                    FontSize =9
                    TabIndex =11
                    BorderColor =11830108
                    Name ="FAID"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =17865
                    LayoutCachedTop =933
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =1173
                    RowStart =3
                    RowEnd =3
                    ColumnStart =19
                    ColumnEnd =20
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =17340
                            Top =933
                            Width =435
                            Height =240
                            Name ="Label1935"
                            Caption ="FAID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17340
                            LayoutCachedTop =933
                            LayoutCachedWidth =17775
                            LayoutCachedHeight =1173
                            RowStart =3
                            RowEnd =3
                            ColumnStart =18
                            ColumnEnd =18
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17865
                    Top =603
                    Width =2415
                    Height =270
                    FontSize =9
                    TabIndex =7
                    BorderColor =11830108
                    Name ="SchNo"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =17865
                    LayoutCachedTop =603
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =873
                    RowStart =2
                    RowEnd =2
                    ColumnStart =19
                    ColumnEnd =20
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =17340
                            Top =603
                            Width =435
                            Height =270
                            Name ="Label1937"
                            Caption ="Sch:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17340
                            LayoutCachedTop =603
                            LayoutCachedWidth =17775
                            LayoutCachedHeight =873
                            RowStart =2
                            RowEnd =2
                            ColumnStart =18
                            ColumnEnd =18
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17865
                    Top =1833
                    Width =1215
                    Height =240
                    FontSize =10
                    TabIndex =23
                    Name ="SubmitSchedule"
                    Caption ="Insert"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17865
                    LayoutCachedTop =1833
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =2073
                    RowStart =6
                    RowEnd =6
                    ColumnStart =19
                    ColumnEnd =19
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17865
                    Top =123
                    Width =2415
                    Height =228
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="AddSchedule"
                    Caption ="Add A  Schedule "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17865
                    LayoutCachedTop =123
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =351
                    ColumnStart =19
                    ColumnEnd =20
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8872517
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =8
                    Left =10833
                    Top =1533
                    Width =2130
                    FontSize =9
                    TabIndex =18
                    BackColor =13952764
                    BorderColor =11830108
                    Name ="MLType"
                    RowSourceType ="Value List"
                    RowSource ="\"1\";\"FA\";\"2\";\"Other\";\"3\";\"RMKT\";\"4\";\"Tracking\";\"5\";\"PortFolio"
                        "\""
                    ColumnWidths ="0;1728"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =255

                    LayoutCachedLeft =10833
                    LayoutCachedTop =1533
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =1773
                    RowStart =5
                    RowEnd =5
                    ColumnStart =14
                    ColumnEnd =15
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =10068
                            Top =1533
                            Width =675
                            Height =240
                            Name ="Label1893"
                            Caption ="MLType:"
                            FontName ="Calibri"
                            LayoutCachedLeft =10068
                            LayoutCachedTop =1533
                            LayoutCachedWidth =10743
                            LayoutCachedHeight =1773
                            RowStart =5
                            RowEnd =5
                            ColumnStart =13
                            ColumnEnd =13
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =9
                            ForeShade =50.0
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
                    Left =13800
                    Top =603
                    Width =3300
                    Height =2925
                    FontSize =9
                    TabIndex =6
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select SchID, SchNo as Schedule, SchFAID As FAID, SchMLRef,DateSchAdded, SchShor"
                        "tDesc FROM Schedule WHERE SchMLRef =216 ORDER BY SchNo;"
                    ColumnWidths ="0;1872;1440;0;0;0"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    AllowValueListEdits =0

                    LayoutCachedLeft =13800
                    LayoutCachedTop =603
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =3528
                    RowStart =2
                    RowEnd =6
                    ColumnStart =17
                    ColumnEnd =17
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17865
                    Top =1233
                    Width =2415
                    FontSize =9
                    TabIndex =14
                    BorderColor =11830108
                    Name ="ShortDesc"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =17865
                    LayoutCachedTop =1233
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =1473
                    RowStart =4
                    RowEnd =4
                    ColumnStart =19
                    ColumnEnd =20
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =17340
                            Top =1233
                            Width =435
                            Height =240
                            Name ="Label2021"
                            Caption ="Desc:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17340
                            LayoutCachedTop =1233
                            LayoutCachedWidth =17775
                            LayoutCachedHeight =1473
                            RowStart =4
                            RowEnd =4
                            ColumnStart =18
                            ColumnEnd =18
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =4770
                    Width =4125
                    Height =240
                    FontSize =9
                    TabIndex =46
                    Name ="SubmitCompany"
                    Caption ="Save / Select Company"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3000
                    LayoutCachedTop =4770
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =5010
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =10
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =8872517
                    ThemeFontIndex =1
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13800
                    Top =3540
                    Width =3300
                    Height =240
                    FontSize =9
                    TabIndex =33
                    Name ="DeleteSchedule"
                    Caption ="Delete  Schedule"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13800
                    LayoutCachedTop =3540
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =3780
                    RowStart =9
                    RowEnd =9
                    ColumnStart =17
                    ColumnEnd =17
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =13800
                    Top =123
                    Width =2235
                    Height =420
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label2106"
                    Caption ="      Select Schedule"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    LayoutCachedLeft =13800
                    LayoutCachedTop =123
                    LayoutCachedWidth =16035
                    LayoutCachedHeight =543
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =20505
                    Top =4305
                    Width =2490
                    FontSize =9
                    TabIndex =43
                    BorderColor =11830108
                    Name ="GroupNo"
                    FontName ="Calibri"

                    LayoutCachedLeft =20505
                    LayoutCachedTop =4305
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =4545
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =3
                    BackThemeColorIndex =4
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =19740
                            Top =4305
                            Width =705
                            Height =240
                            Name ="Label2126"
                            Caption ="Group:"
                            FontName ="Calibri"
                            LayoutCachedLeft =19740
                            LayoutCachedTop =4305
                            LayoutCachedWidth =20445
                            LayoutCachedHeight =4545
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =20505
                    Top =5235
                    Width =2490
                    FontSize =9
                    TabIndex =51
                    BorderColor =11830108
                    Name ="GrpSchID"
                    FontName ="Calibri"

                    LayoutCachedLeft =20505
                    LayoutCachedTop =5235
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =5475
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =3
                    BackThemeColorIndex =4
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =19740
                            Top =5235
                            Width =705
                            Height =240
                            Name ="Label2128"
                            Caption ="SchID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =19740
                            LayoutCachedTop =5235
                            LayoutCachedWidth =20445
                            LayoutCachedHeight =5475
                            RowStart =5
                            RowEnd =5
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =20505
                    Top =5535
                    Width =1350
                    Height =240
                    FontSize =9
                    TabIndex =53
                    Name ="SubmitGroup"
                    Caption ="Insert"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20505
                    LayoutCachedTop =5535
                    LayoutCachedWidth =21855
                    LayoutCachedHeight =5775
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =20505
                    Top =3900
                    Width =2490
                    Height =240
                    FontSize =10
                    FontWeight =700
                    TabIndex =36
                    Name ="AddGroup"
                    Caption ="Add A Unit Group "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20505
                    LayoutCachedTop =3900
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =4140
                    ColumnStart =2
                    ColumnEnd =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8872517
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =16200
                    Top =4305
                    Width =3060
                    Height =1785
                    FontSize =9
                    TabIndex =42
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select SGrpID, SGrpNO As GroupNo, LeaseType As GType, SGrpFAID, SgrpDesc, SGrpSc"
                        "hID FROM SchGrp WHERE SgrpSchID = 893 ORDER BY SGrpNO;"
                    ColumnWidths ="0;1728;1152;0;0;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =16200
                    LayoutCachedTop =4305
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =6090
                    RowStart =2
                    RowEnd =7
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =20505
                    Top =4935
                    Width =2490
                    FontSize =9
                    TabIndex =48
                    BorderColor =11830108
                    Name ="GroupDesc"
                    FontName ="Calibri"

                    LayoutCachedLeft =20505
                    LayoutCachedTop =4935
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =5175
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =3
                    BackThemeColorIndex =4
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =19740
                            Top =4935
                            Width =705
                            Height =240
                            Name ="Label2135"
                            Caption ="Desc:"
                            FontName ="Calibri"
                            LayoutCachedLeft =19740
                            LayoutCachedTop =4935
                            LayoutCachedWidth =20445
                            LayoutCachedHeight =5175
                            RowStart =4
                            RowEnd =4
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16200
                    Top =6150
                    Width =3060
                    Height =240
                    FontSize =9
                    TabIndex =56
                    Name ="DeleteGroup"
                    Caption ="Delete  Group"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =16200
                    LayoutCachedTop =6150
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =6390
                    RowStart =8
                    RowEnd =8
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =16200
                    Top =3900
                    Width =3060
                    Height =345
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label2137"
                    Caption ="      Select Target Group"
                    FontName ="Calibri"
                    LayoutCachedLeft =16200
                    LayoutCachedTop =3900
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =4245
                    RowEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =20505
                    Top =4605
                    Width =2490
                    Height =270
                    FontSize =9
                    TabIndex =45
                    BorderColor =11830108
                    Name ="GroupType"
                    RowSourceType ="Value List"
                    RowSource ="\"FMV\";T.R.A.C.;SPLIT T.R.A.C.;CAPITAL;LOAN;REMARKET;REPORTING"
                    FontName ="Calibri"

                    LayoutCachedLeft =20505
                    LayoutCachedTop =4605
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =4875
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =19740
                            Top =4605
                            Width =705
                            Height =270
                            Name ="Label2124"
                            Caption ="Type:"
                            FontName ="Calibri"
                            LayoutCachedLeft =19740
                            LayoutCachedTop =4605
                            LayoutCachedWidth =20445
                            LayoutCachedHeight =4875
                            RowStart =3
                            RowEnd =3
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =5835
                    Top =7095
                    Width =7905
                    Height =7065
                    FontSize =8
                    TabIndex =80
                    Name ="UnitList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FleetStudyImport.VIN, FleetStudyImport.UnitNumber, FleetStudyImport.UnitY"
                        "ear, FleetStudyImport.UnitMake, FleetStudyImport.UnitModel, [SubType] & \" \" & "
                        "[Type] AS ClassType, FleetStudyImport.Location, FleetStudyImport.Division, Fleet"
                        "StudyImport.DivisionID FROM FleetStudyImport ORDER BY FleetStudyImport.Location,"
                        " FleetStudyImport.SubType, FleetStudyImport.UnitYear, FleetStudyImport.UnitMake;"
                    ColumnWidths ="1728;0;576;1008;864;864;1440;1296;0"
                    ControlTipText ="Unit to be moved"
                    VerticalAnchor =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =7095
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =14160
                    RowStart =2
                    RowEnd =6
                    ColumnEnd =4
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =16185
                    Top =6915
                    Width =6855
                    Height =7245
                    FontSize =8
                    TabIndex =79
                    BackColor =15590879
                    Name ="UnitsFormsList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="864;2160;720;720;720;720"
                    ControlTipText ="Units in Target Group"
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =16185
                    LayoutCachedTop =6915
                    LayoutCachedWidth =23040
                    LayoutCachedHeight =14160
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11190
                    Top =6300
                    Width =1230
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =63
                    Name ="Text2215"
                    ControlSource ="=\"Units: \" & [UnitList].[ListCount]"
                    FontName ="Calibri"
                    BottomPadding =0

                    LayoutCachedLeft =11190
                    LayoutCachedTop =6300
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =6615
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19365
                    Top =6540
                    Width =1005
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =65
                    Name ="Text2216"
                    ControlSource ="=\"Units: \" & [UnitsFormsList].[ListCount]-1"
                    FontName ="Calibri"

                    LayoutCachedLeft =19365
                    LayoutCachedTop =6540
                    LayoutCachedWidth =20370
                    LayoutCachedHeight =6855
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5850
                    Top =6300
                    Width =5280
                    Height =315
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label2217"
                    Caption ="     Select Units To Move"
                    FontName ="Calibri"
                    BottomPadding =0
                    LayoutCachedLeft =5850
                    LayoutCachedTop =6300
                    LayoutCachedWidth =11130
                    LayoutCachedHeight =6615
                    ColumnEnd =4
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14400
                    Top =9000
                    Height =240
                    FontSize =9
                    TabIndex =83
                    Name ="MoveUnits"
                    Caption ="Move ► "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =14400
                    LayoutCachedTop =9000
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =9240
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14415
                    Top =8655
                    Height =240
                    FontSize =9
                    TabIndex =82
                    Name ="GetAll"
                    Caption ="▼ Select All Units ▼"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14415
                    LayoutCachedTop =8655
                    LayoutCachedWidth =15855
                    LayoutCachedHeight =8895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12645
                    Top =4605
                    Width =1320
                    Height =270
                    FontSize =9
                    TabIndex =44
                    BorderColor =11830108
                    Name ="LocCounty"
                    FontName ="Calibri"
                    LeftPadding =60

                    LayoutCachedLeft =12645
                    LayoutCachedTop =4605
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =4875
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =11700
                            Top =4605
                            Width =855
                            Height =270
                            FontWeight =700
                            Name ="Label2227"
                            Caption ="County:"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =11700
                            LayoutCachedTop =4605
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =4875
                            RowStart =6
                            RowEnd =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12645
                    Top =4005
                    Width =3060
                    FontSize =9
                    TabIndex =37
                    BorderColor =11830108
                    Name ="LocName"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =12645
                    LayoutCachedTop =4005
                    LayoutCachedWidth =15705
                    LayoutCachedHeight =4245
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =11700
                            Top =4005
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="Label2229"
                            Caption ="Location:"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =11700
                            LayoutCachedTop =4005
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =4245
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12645
                    Top =4935
                    Width =1320
                    FontSize =9
                    TabIndex =47
                    BorderColor =11830108
                    Name ="LocPhone"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    LeftPadding =60

                    LayoutCachedLeft =12645
                    LayoutCachedTop =4935
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =5175
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =11700
                            Top =4935
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="Label2231"
                            Caption ="Phone:"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =11700
                            LayoutCachedTop =4935
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =5175
                            RowStart =7
                            RowEnd =7
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12645
                    Top =5535
                    Width =3060
                    Height =240
                    FontSize =9
                    TabIndex =52
                    Name ="SubmitLocation"
                    Caption ="Submit Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12645
                    LayoutCachedTop =5535
                    LayoutCachedWidth =15705
                    LayoutCachedHeight =5775
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =8
                    HoverTint =60.0
                    PressedThemeColorIndex =8
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9795
                    Top =2460
                    Width =1650
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =28
                    Name ="AddLocation"
                    Caption ="Add A Location "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    RightPadding =45
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9795
                    LayoutCachedTop =2460
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =2760
                    ColumnStart =2
                    ColumnEnd =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8872517
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =7530
                    Top =2835
                    Width =3915
                    Height =3105
                    FontSize =9
                    TabIndex =29
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationType, Locations.LocationName AS L"
                        "ocName, Locations.LCity AS City, Locations.LState AS State, Locations.Locationsn"
                        " AS LOC, Locations.LCounty, Locations.LCountry, Locations.LPhone, Locations.Clie"
                        "ntRef FROM Locations WHERE (((Locations.[ClientRef])=5215) AND ((Locations.[Loca"
                        "tionType])='Garage')) ORDER BY Locations.LocationName;"
                    ColumnWidths ="0;0;1440;1152;432;720;0;0;0;0"
                    OnClick ="[Event Procedure]"
                    RightPadding =45
                    AllowValueListEdits =0

                    LayoutCachedLeft =7530
                    LayoutCachedTop =2835
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =5940
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12645
                    Top =5235
                    Width =1320
                    FontSize =9
                    TabIndex =49
                    BorderColor =11830108
                    Name ="LocCountry"
                    FontName ="Calibri"
                    LeftPadding =60

                    LayoutCachedLeft =12645
                    LayoutCachedTop =5235
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =5475
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =11700
                            Top =5235
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="Label2238"
                            Caption ="Country:"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =11700
                            LayoutCachedTop =5235
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =5475
                            RowStart =8
                            RowEnd =8
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7530
                    Top =5940
                    Width =3915
                    Height =240
                    FontSize =9
                    TabIndex =55
                    Name ="DeleteLocation"
                    Caption ="Delete  Selected Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =7530
                    LayoutCachedTop =5940
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =6180
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7530
                    Top =2460
                    Width =2205
                    Height =300
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label2240"
                    Caption ="     Select Target Location"
                    FontName ="Calibri"
                    BottomPadding =45
                    LayoutCachedLeft =7530
                    LayoutCachedTop =2460
                    LayoutCachedWidth =9735
                    LayoutCachedHeight =2760
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12645
                    Top =4305
                    Width =1320
                    FontSize =9
                    TabIndex =39
                    BorderColor =11830108
                    Name ="LocCity"
                    FontName ="Calibri"
                    LeftPadding =60

                    LayoutCachedLeft =12645
                    LayoutCachedTop =4305
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =4545
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =11700
                            Top =4305
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="Label2289"
                            Caption ="City/St/Zip:"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =11700
                            LayoutCachedTop =4305
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =4545
                            RowStart =5
                            RowEnd =5
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14055
                    Top =4305
                    Width =645
                    FontSize =9
                    TabIndex =40
                    BorderColor =11830108
                    Name ="LocState"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_States"
                    StatusBarText ="lookup states table"
                    FontName ="Calibri"
                    LeftPadding =60

                    LayoutCachedLeft =14055
                    LayoutCachedTop =4305
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =4545
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14760
                    Top =4305
                    Width =945
                    FontSize =9
                    TabIndex =41
                    BorderColor =11830108
                    Name ="LocZip"
                    FontName ="Calibri"
                    InputMask ="00000\\-9999;;_"
                    RightPadding =75

                    LayoutCachedLeft =14760
                    LayoutCachedTop =4305
                    LayoutCachedWidth =15705
                    LayoutCachedHeight =4545
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16185
                    Top =6540
                    Width =3093
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =64
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =13020235
                    Name ="Label2219"
                    ControlSource ="=[GroupNo] & \" | \" & [GroupDesc]"
                    FontName ="Calibri"

                    LayoutCachedLeft =16185
                    LayoutCachedTop =6540
                    LayoutCachedWidth =19278
                    LayoutCachedHeight =6855
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =2880
                    Top =72
                    Width =259
                    Height =300
                    FontSize =14
                    Name ="Label2467"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =72
                    LayoutCachedWidth =3139
                    LayoutCachedHeight =372
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =7320
                    Top =120
                    Width =259
                    Height =330
                    FontSize =14
                    Name ="Label2468"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =7579
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =13560
                    Top =120
                    Width =259
                    Height =330
                    FontSize =14
                    Name ="Label2469"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =13560
                    LayoutCachedTop =120
                    LayoutCachedWidth =13819
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =7140
                    Top =2460
                    Width =330
                    Height =330
                    FontSize =14
                    Name ="Label2470"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =7140
                    LayoutCachedTop =2460
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =2790
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =15960
                    Top =3900
                    Width =330
                    Height =330
                    FontSize =14
                    Name ="Label2472"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =15960
                    LayoutCachedTop =3900
                    LayoutCachedWidth =16290
                    LayoutCachedHeight =4230
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =5820
                    Top =6301
                    Width =330
                    Height =330
                    FontSize =14
                    Name ="Label2473"
                    Caption =""
                    FontName ="Wingdings"
                    LayoutCachedLeft =5820
                    LayoutCachedTop =6301
                    LayoutCachedWidth =6150
                    LayoutCachedHeight =6631
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6180
                    Top =1860
                    Width =975
                    FontSize =9
                    TabIndex =20
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupId"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"
                    BottomPadding =120

                    LayoutCachedLeft =6180
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =2100
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5430
                            Top =1860
                            Width =690
                            Height =240
                            Name ="Label331"
                            Caption =" Group:"
                            FontName ="Calibri"
                            BottomPadding =120
                            LayoutCachedLeft =5430
                            LayoutCachedTop =1860
                            LayoutCachedWidth =6120
                            LayoutCachedHeight =2100
                            RowStart =6
                            RowEnd =6
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =6615
                    Width =1725
                    FontSize =9
                    TabIndex =66
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchVin"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search VIN"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =6615
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =6855
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10875
                    Top =6615
                    FontSize =9
                    TabIndex =70
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchLoc"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search Location"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0

                    LayoutCachedLeft =10875
                    LayoutCachedTop =6615
                    LayoutCachedWidth =12315
                    LayoutCachedHeight =6855
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14925
                    Top =5235
                    Width =735
                    FontSize =9
                    TabIndex =50
                    Name ="LocationId"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"
                    LeftPadding =60

                    LayoutCachedLeft =14925
                    LayoutCachedTop =5235
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =5475
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =13980
                            Top =5235
                            Width =855
                            Height =240
                            Name ="Label2714"
                            Caption ="LocationId"
                            FontName ="Calibri"
                            LeftPadding =60
                            LayoutCachedLeft =13980
                            LayoutCachedTop =5235
                            LayoutCachedWidth =14835
                            LayoutCachedHeight =5475
                            RowStart =10
                            RowEnd =10
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =21915
                    Top =5535
                    Width =1080
                    Height =240
                    FontSize =9
                    TabIndex =54
                    Name ="UpdateGroup"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =21915
                    LayoutCachedTop =5535
                    LayoutCachedWidth =22995
                    LayoutCachedHeight =5775
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14370
                    Top =11520
                    Height =240
                    FontSize =9
                    TabIndex =85
                    Name ="DeleteUnits"
                    Caption ="Delete Selected►"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =14370
                    LayoutCachedTop =11520
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =11760
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14370
                    Top =11040
                    Height =240
                    FontSize =9
                    TabIndex =84
                    Name ="SelectToDelete"
                    Caption ="Select All Units ►"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14370
                    LayoutCachedTop =11040
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =11280
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12033
                    Top =1833
                    Width =930
                    Height =240
                    FontSize =9
                    TabIndex =22
                    Name ="UpdateMstLease"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12033
                    LayoutCachedTop =1833
                    LayoutCachedWidth =12963
                    LayoutCachedHeight =2073
                    RowStart =6
                    RowEnd =6
                    ColumnStart =15
                    ColumnEnd =15
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =6855
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =58
                    Name ="UpdateModel"
                    Caption ="Update model \015\012(after VIN decode)"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =6855
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =7335
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =7410
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =59
                    Name ="UnitNumToObcNum"
                    Caption ="Update \015\012UnitOBCNumber"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =7410
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =7890
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =6300
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =57
                    Name ="PostImport"
                    Caption ="Set  Group \015\012Unit Count\015\012 "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =6300
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =6780
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7560
                    Top =6615
                    Width =570
                    FontSize =9
                    TabIndex =67
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchYear"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search Year"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7560
                    LayoutCachedTop =6615
                    LayoutCachedWidth =8130
                    LayoutCachedHeight =6855
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10005
                    Top =6615
                    Width =870
                    FontSize =9
                    TabIndex =69
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchType"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search Type"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10005
                    LayoutCachedTop =6615
                    LayoutCachedWidth =10875
                    LayoutCachedHeight =6855
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8130
                    Top =6615
                    Width =1008
                    FontSize =9
                    TabIndex =68
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchMake"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search Make"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8130
                    LayoutCachedTop =6615
                    LayoutCachedWidth =9138
                    LayoutCachedHeight =6855
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin EmptyCell
                    Left =9135
                    Top =6615
                    Width =864
                    Name ="EmptyCell2790"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9135
                    LayoutCachedTop =6615
                    LayoutCachedWidth =9999
                    LayoutCachedHeight =6855
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =45
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =7965
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =60
                    Name ="GrabVINStoDecode"
                    Caption ="Grab VINS to Decode"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =7965
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =8445
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =8520
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =61
                    Name ="CheckVINsOnly"
                    Caption ="Decode 'grabbed' VINs"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =8520
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =9000
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =19215
                    Top =1833
                    Width =1065
                    Height =240
                    FontSize =9
                    TabIndex =24
                    Name ="UpdateSchedule"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19215
                    LayoutCachedTop =1833
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =2073
                    RowStart =6
                    RowEnd =6
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-117
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =10
                    ListWidth =2448
                    Left =14175
                    Top =7155
                    Width =1725
                    Height =300
                    FontSize =10
                    TabIndex =81
                    BackColor =15523798
                    Name ="DivisionOwnerID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientDivisions.DivisionID, ClientDivisions.DivisionName, ClientDivisions"
                        ".ClientID FROM ClientDivisions WHERE (((ClientDivisions.ClientID)=[Forms]![Units"
                        "Form].[clientID]));"
                    ColumnWidths ="0;2304"
                    FontName ="Calibri"
                    Tag ="Cloned"
                    ControlTipText ="The Division or Owership Group This Unit belongs to."
                    AllowValueListEdits =1
                    InheritValueList =1

                    LayoutCachedLeft =14175
                    LayoutCachedTop =7155
                    LayoutCachedWidth =15900
                    LayoutCachedHeight =7455
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14175
                            Top =6885
                            Width =1725
                            Height =240
                            Name ="Label4682"
                            Caption ="Division To Use"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =14175
                            LayoutCachedTop =6885
                            LayoutCachedWidth =15900
                            LayoutCachedHeight =7125
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =8130
                    Top =6855
                    Width =1008
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =74
                    ForeColor =8210719
                    Name ="ToggleSortMake1"
                    DefaultValue ="0"
                    Caption ="Make"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =8130
                    LayoutCachedTop =6855
                    LayoutCachedWidth =9138
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7560
                    Top =6855
                    Width =570
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =73
                    ForeColor =8210719
                    Name ="ToggleSortYear1"
                    DefaultValue ="0"
                    Caption ="Year"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =7560
                    LayoutCachedTop =6855
                    LayoutCachedWidth =8130
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =10005
                    Top =6855
                    Width =870
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =76
                    ForeColor =8210719
                    Name ="ToggleSortClassType1"
                    DefaultValue ="0"
                    Caption ="Class Type"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =10005
                    LayoutCachedTop =6855
                    LayoutCachedWidth =10875
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =10875
                    Top =6855
                    Width =1440
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =77
                    ForeColor =8210719
                    Name ="ToggleSortLocation1"
                    DefaultValue ="0"
                    Caption ="Location"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =10875
                    LayoutCachedTop =6855
                    LayoutCachedWidth =12315
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5835
                    Top =6855
                    Width =1725
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =72
                    ForeColor =8210719
                    Name ="ToggleSortVIN1"
                    DefaultValue ="0"
                    Caption ="VIN"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =5835
                    LayoutCachedTop =6855
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =9135
                    Top =6855
                    Width =864
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =75
                    ForeColor =8210719
                    Name ="ToggleSortModel1"
                    DefaultValue ="0"
                    Caption ="Model"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =9135
                    LayoutCachedTop =6855
                    LayoutCachedWidth =9999
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =12330
                    Top =6855
                    Width =1410
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =78
                    ForeColor =8210719
                    Name ="ToggleSortDivision1"
                    DefaultValue ="0"
                    Caption ="Division"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =12330
                    LayoutCachedTop =6855
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =7095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =45
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12330
                    Top =6615
                    Width =1410
                    FontSize =9
                    TabIndex =71
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchDivision"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Search Location"
                    GroupTable =45
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0

                    LayoutCachedLeft =12330
                    LayoutCachedTop =6615
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =6855
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =45
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3060
                    Top =9075
                    Width =2040
                    Height =480
                    FontSize =10
                    TabIndex =62
                    Name ="SyncLocationDivisions"
                    Caption ="Sync Location/Division"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =48
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3060
                    LayoutCachedTop =9075
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =9555
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    GroupTable =48
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2895
                    Top =1875
                    Width =840
                    Height =240
                    Name ="Label3097"
                    Caption =" Corp State:"
                    FontName ="Calibri"
                    BottomPadding =120
                    LayoutCachedLeft =2895
                    LayoutCachedTop =1875
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =2115
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "FleetStudy_Importer.cls"
