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
    Width =23040
    DatasheetFontHeight =10
    ItemSuffix =2275
    Left =-1320
    Top =4740
    Right =18405
    Bottom =15075
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xcdd0dbf18d70e540
    End
    RecordSource ="Select * From Clients Where ClientID=0"
    Caption ="Clients"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            Height =720
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =21465
                    Top =180
                    Width =555
                    Height =228
                    TabIndex =7
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

                    LayoutCachedLeft =21465
                    LayoutCachedTop =180
                    LayoutCachedWidth =22020
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
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =615
                    Top =45
                    Width =9480
                    Height =540
                    ColumnOrder =5
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="ClientName"
                    ControlSource ="clientCompanyName"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =615
                    LayoutCachedTop =45
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =585
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =540
                    Width =23040
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =23040
                    LayoutCachedHeight =720
                    TabIndex =9
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =14540
                    Top =180
                    Width =1605
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="AddNewCompany"
                    Caption ="Add a Company"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14540
                    LayoutCachedTop =180
                    LayoutCachedWidth =16145
                    LayoutCachedHeight =468
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    ForeShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    HoverForeShade =95.0
                    PressedForeThemeColorIndex =1
                    PressedForeShade =95.0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19200
                    Top =195
                    Width =780
                    Height =210
                    ColumnOrder =3
                    TabIndex =3
                    Name ="ClientID"
                    ControlSource ="clientID"
                    FontName ="Calibri"

                    LayoutCachedLeft =19200
                    LayoutCachedTop =195
                    LayoutCachedWidth =19980
                    LayoutCachedHeight =405
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =18525
                            Top =195
                            Width =645
                            Height =210
                            Name ="Label641"
                            Caption ="ClientID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =18525
                            LayoutCachedTop =195
                            LayoutCachedWidth =19170
                            LayoutCachedHeight =405
                            ThemeFontIndex =1
                            BackThemeColorIndex =6
                            BackShade =75.0
                            BorderThemeColorIndex =1
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            ForeShade =50.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =18870
                    Width =1650
                    ColumnOrder =4
                    Name ="nmch"
                    FontName ="Calibri"
                    InputMask ="Password"

                    LayoutCachedLeft =18870
                    LayoutCachedWidth =20520
                    LayoutCachedHeight =240
                    ThemeFontIndex =1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            FontItalic = NotDefault
                            OverlapFlags =85
                            Left =17745
                            Width =1065
                            Height =165
                            Name ="Label453"
                            Caption ="NM Change PW:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17745
                            LayoutCachedWidth =18810
                            LayoutCachedHeight =165
                            ThemeFontIndex =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20640
                    Top =195
                    Width =585
                    Height =210
                    ColumnOrder =0
                    TabIndex =4
                    Name ="ClientGroupID2"
                    ControlSource ="ClientGroupID"
                    FontName ="Calibri"

                    LayoutCachedLeft =20640
                    LayoutCachedTop =195
                    LayoutCachedWidth =21225
                    LayoutCachedHeight =405
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =1
                            Left =19965
                            Top =195
                            Width =645
                            Height =210
                            Name ="Label1873"
                            Caption ="CGroup"
                            FontName ="Calibri"
                            LayoutCachedLeft =19965
                            LayoutCachedTop =195
                            LayoutCachedWidth =20610
                            LayoutCachedHeight =405
                            ThemeFontIndex =1
                            BackThemeColorIndex =6
                            BackShade =75.0
                            BorderThemeColorIndex =1
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            ForeShade =50.0
                        End
                    End
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureAlignment =0
                    PictureType =2
                    Left =60
                    Width =541
                    Height =645
                    Name ="Image1161"
                    Picture ="ATLAASGlobeLogo"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =601
                    LayoutCachedHeight =645
                    TabIndex =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12360
                    Top =180
                    Width =2135
                    Height =288
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="ImportFAID"
                    Caption ="Import Salesforce FAID "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =90
                    TopPadding =45
                    RightPadding =105
                    BottomPadding =150

                    LayoutCachedLeft =12360
                    LayoutCachedTop =180
                    LayoutCachedWidth =14495
                    LayoutCachedHeight =468
                    ForeThemeColorIndex =1
                    ForeShade =95.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =2
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =22032
                    Top =180
                    Width =750
                    Height =228
                    TabIndex =6
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =22032
                    LayoutCachedTop =180
                    LayoutCachedWidth =22782
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11865
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =13125
                    Top =2025
                    Width =1440
                    Height =480
                    Name ="Box2"
                    LayoutCachedLeft =13125
                    LayoutCachedTop =2025
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =2505
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =14475
                    Top =2130
                    Width =240
                    Height =315
                    Name ="Box1"
                    LayoutCachedLeft =14475
                    LayoutCachedTop =2130
                    LayoutCachedWidth =14715
                    LayoutCachedHeight =2445
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =12840
                    Top =2190
                    Width =420
                    Height =300
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =2190
                    LayoutCachedWidth =13260
                    LayoutCachedHeight =2490
                    TabIndex =35
                End
                Begin Tab
                    OverlapFlags =85
                    Left =2893
                    Top =3150
                    Width =20087
                    Height =8715
                    TabIndex =32
                    Name ="TabControls"
                    VerticalAnchor =2

                    LayoutCachedLeft =2893
                    LayoutCachedTop =3150
                    LayoutCachedWidth =22980
                    LayoutCachedHeight =11865
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="Schedules"
                            OnClick ="[Event Procedure]"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3030
                                    Top =3600
                                    Width =19560
                                    Height =7890
                                    Name ="ClientsView_Schedules1"
                                    SourceObject ="Form.ClientsView_Schedules1"
                                    LinkChildFields ="Client"
                                    LinkMasterFields ="clientID"

                                    LayoutCachedLeft =3030
                                    LayoutCachedTop =3600
                                    LayoutCachedWidth =22590
                                    LayoutCachedHeight =11490
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="Contacts"
                            Caption ="Contacts"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3090
                                    Top =3630
                                    Width =17235
                                    Height =8100
                                    BorderColor =967423
                                    Name ="ClientContact_sub"
                                    SourceObject ="Form.ClientContact_sub"
                                    LinkChildFields ="ClientId"
                                    LinkMasterFields ="clientID"

                                    LayoutCachedLeft =3090
                                    LayoutCachedTop =3630
                                    LayoutCachedWidth =20325
                                    LayoutCachedHeight =11730
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="Locations | Divisions"
                            EventProcPrefix ="Locations___Divisions"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3045
                                    Top =3630
                                    Width =19800
                                    Height =7905
                                    Name ="Locations_sub"
                                    SourceObject ="Form.Locations_sub"
                                    LinkChildFields ="ClientGroupID"
                                    LinkMasterFields ="ClientGroupID"

                                    LayoutCachedLeft =3045
                                    LayoutCachedTop =3630
                                    LayoutCachedWidth =22845
                                    LayoutCachedHeight =11535
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="AwardedDeals"
                            Caption ="Awarded Deals"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3090
                                    Top =3630
                                    Width =19726
                                    Height =8025
                                    Name ="ProposalsSummary_sub"
                                    SourceObject ="Form.ProposalsSummary_sub"
                                    LinkChildFields ="clientRef"
                                    LinkMasterFields ="clientID"

                                    LayoutCachedLeft =3090
                                    LayoutCachedTop =3630
                                    LayoutCachedWidth =22816
                                    LayoutCachedHeight =11655
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="Notes tab"
                            EventProcPrefix ="Notes_tab"
                            Caption ="Notes"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="Master Lease"
                            EventProcPrefix ="Master_Lease"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3120
                                    Top =3645
                                    Width =19305
                                    Height =8010
                                    Name ="Leasing_sub"
                                    SourceObject ="Form.Leasing_sub"
                                    LinkChildFields ="Client"
                                    LinkMasterFields ="clientID"

                                    LayoutCachedLeft =3120
                                    LayoutCachedTop =3645
                                    LayoutCachedWidth =22425
                                    LayoutCachedHeight =11655
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="ValueAdded"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3112
                                    Top =3705
                                    Width =19304
                                    Height =7830
                                    Name ="ValueAdded_sub"
                                    SourceObject ="Form.ValueAdded"
                                    LinkChildFields ="CONCLIENTREF"
                                    LinkMasterFields ="clientID"

                                    LayoutCachedLeft =3112
                                    LayoutCachedTop =3705
                                    LayoutCachedWidth =22416
                                    LayoutCachedHeight =11535
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3030
                            Top =3555
                            Width =19815
                            Height =8175
                            Name ="SUFConditions"
                            Caption ="SUF Conditions"
                            LayoutCachedLeft =3030
                            LayoutCachedTop =3555
                            LayoutCachedWidth =22845
                            LayoutCachedHeight =11730
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =3030
                                    Top =3585
                                    Width =19545
                                    Height =8040
                                    Name ="SUF_Conditions"
                                    SourceObject ="Form.SUFConditions"
                                    LinkChildFields ="ClientID"
                                    LinkMasterFields ="ClientID"

                                    LayoutCachedLeft =3030
                                    LayoutCachedTop =3585
                                    LayoutCachedWidth =22575
                                    LayoutCachedHeight =11625
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =2895
                    Top =45
                    Width =4830
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label120"
                    Caption ="Company Details"
                    FontName ="Calibri"
                    LayoutCachedLeft =2895
                    LayoutCachedTop =45
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =285
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =345
                    Width =3675
                    Height =270
                    FontSize =9
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =967423
                    Name ="clientCompanyName"
                    ControlSource ="clientCompanyName"
                    ValidationText ="Client Name is Required"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Required - The 'Proper' Name of the Client/Company"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =345
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =615
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =345
                            Width =1215
                            Height =270
                            BorderColor =967423
                            Name ="Label138"
                            Caption ="Company Name:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =345
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =615
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =975
                    Width =3675
                    FontSize =9
                    TabIndex =7
                    Name ="clientDBA"
                    ControlSource ="clientDBA"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =975
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =1215
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =975
                            Width =1215
                            Height =240
                            Name ="Label139"
                            Caption ="d/b/a:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =975
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1215
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =1575
                    Width =3675
                    FontSize =9
                    TabIndex =12
                    BorderColor =11830108
                    Name ="clientAddress1"
                    ControlSource ="clientAddress1"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1575
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =1815
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =1575
                            Width =1215
                            Height =240
                            Name ="Label140"
                            Caption ="Address1:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =1575
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1815
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =1875
                    Width =3675
                    FontSize =9
                    TabIndex =14
                    BorderColor =11830108
                    Name ="clientAddress2"
                    ControlSource ="clientAddress2"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1875
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =2115
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =1875
                            Width =1215
                            Height =240
                            Name ="Label141"
                            Caption ="Address2:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =1875
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =2115
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =2175
                    Width =1770
                    FontSize =9
                    TabIndex =16
                    BorderColor =11830108
                    Name ="clientCity"
                    ControlSource ="clientCity"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2175
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =2415
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =2175
                            Width =1215
                            Height =240
                            Name ="Label142"
                            Caption ="City/State/Zip:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =2175
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =2415
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6030
                    Top =2175
                    Width =705
                    FontSize =9
                    TabIndex =17
                    BorderColor =11830108
                    Name ="clientState"
                    ControlSource ="clientState"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_States"
                    StatusBarText ="lookup states table"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =6030
                    LayoutCachedTop =2175
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =2415
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6795
                    Top =2175
                    Width =1080
                    FontSize =9
                    TabIndex =18
                    BorderColor =11830108
                    Name ="clientZip"
                    ControlSource ="clientZip"
                    FontName ="Calibri"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =2

                    LayoutCachedLeft =6795
                    LayoutCachedTop =2175
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =2415
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =975
                    Width =1635
                    FontSize =9
                    TabIndex =8
                    Name ="clientMainPhone"
                    ControlSource ="clientMainPhone"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =975
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =1215
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =975
                            Width =950
                            Height =240
                            Name ="Label27"
                            Caption ="Main Phone:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =975
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =1215
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =1275
                    Width =1635
                    FontSize =9
                    TabIndex =10
                    Name ="clientFax"
                    ControlSource ="clientFax"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =1275
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =1275
                            Width =950
                            Height =240
                            Name ="Label31"
                            Caption ="Fax Number:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =1275
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =1515
                            RowStart =3
                            RowEnd =3
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
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
                    Left =8940
                    Top =1575
                    Width =3930
                    FontSize =9
                    TabIndex =13
                    Name ="clientWebsite"
                    ControlSource ="clientWebsite"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =1575
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =1815
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =1575
                            Width =950
                            Height =240
                            Name ="Label35"
                            Caption ="Website:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =1575
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =1815
                            RowStart =4
                            RowEnd =4
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14220
                    Top =450
                    Width =1905
                    Height =270
                    FontSize =9
                    TabIndex =24
                    Name ="clientInitContact"
                    ControlSource ="clientInitContact"
                    Format ="Short Date"
                    FontName ="Calibri"

                    LayoutCachedLeft =14220
                    LayoutCachedTop =450
                    LayoutCachedWidth =16125
                    LayoutCachedHeight =720
                    ColumnStart =13
                    ColumnEnd =13
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =13035
                            Top =450
                            Width =1125
                            Height =270
                            Name ="Label187"
                            Caption ="Initial Contact:"
                            FontName ="Calibri"
                            LayoutCachedLeft =13035
                            LayoutCachedTop =450
                            LayoutCachedWidth =14160
                            LayoutCachedHeight =720
                            ColumnStart =12
                            ColumnEnd =12
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
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14220
                    Top =780
                    Width =1905
                    Height =270
                    FontSize =9
                    TabIndex =26
                    Name ="clientBecameDate"
                    ControlSource ="clientBecameDate"
                    Format ="Short Date"
                    FontName ="Calibri"

                    LayoutCachedLeft =14220
                    LayoutCachedTop =780
                    LayoutCachedWidth =16125
                    LayoutCachedHeight =1050
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =13020
                            Top =780
                            Width =1140
                            Height =270
                            Name ="Label43"
                            Caption ="Became Client:"
                            FontName ="Calibri"
                            LayoutCachedLeft =13020
                            LayoutCachedTop =780
                            LayoutCachedWidth =14160
                            LayoutCachedHeight =1050
                            RowStart =1
                            RowEnd =1
                            ColumnStart =12
                            ColumnEnd =12
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
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4200
                    Top =2775
                    Width =1770
                    Height =255
                    FontSize =9
                    TabIndex =23
                    Name ="LeaseSTofCorp"
                    ControlSource ="LeaseSTofCorp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev, lookup_States.StateLong FROM lookup_States ORD"
                        "ER BY lookup_States.StateAbbrev;"
                    ColumnWidths ="0;1728"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2775
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =3030
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =2775
                            Width =1215
                            Height =255
                            Name ="Label249"
                            Caption ="Incorporated In:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =2775
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =3030
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =1275
                    Width =3675
                    FontSize =9
                    TabIndex =9
                    Name ="clientTaxID"
                    ControlSource ="clientTaxID"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1275
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =1275
                            Width =1215
                            Height =240
                            Name ="Label248"
                            Caption ="Tax ID:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =1275
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1515
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =4200
                    Top =2475
                    Width =1770
                    FontSize =9
                    TabIndex =20
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientGroups.ClientGroupId, ClientGroups.ClientGroupName, ClientGroups.Is"
                        "Deleted FROM ClientGroups WHERE (((ClientGroups.IsDeleted)=False)) ORDER BY Clie"
                        "ntGroups.ClientGroupName;"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2475
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =2715
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =2475
                            Width =1215
                            Height =240
                            Name ="Label331"
                            Caption ="Client Group:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =2475
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =2715
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
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
                    Left =6795
                    Top =2475
                    Width =1080
                    ColumnWidth =2535
                    FontSize =9
                    TabIndex =21
                    Name ="CLIENTGRPCNTRY"
                    ControlSource ="CLIENTGRPCNTRY"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Country.CNID, Country.Country FROM Country;"
                    ColumnWidths ="0;1440"
                    FontName ="Calibri"
                    GroupTable =2
                    AllowValueListEdits =1

                    LayoutCachedLeft =6795
                    LayoutCachedTop =2475
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =2715
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =6030
                            Top =2475
                            Width =705
                            Height =240
                            Name ="Label333"
                            Caption ="Foreign?:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =6030
                            LayoutCachedTop =2475
                            LayoutCachedWidth =6735
                            LayoutCachedHeight =2715
                            RowStart =7
                            RowEnd =7
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8940
                    Top =345
                    Width =1635
                    Height =270
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =967423
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    StatusBarText ="Required - 'Short' Name for this Client, This Name will be used in many reports."
                    ValidationRule ="Is Not Null"
                    ValidationText ="Short Name is required"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Required - 'Short' Name for this Client, This Name will be used in many reports."
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =345
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =615
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =7935
                            Top =345
                            Width =950
                            Height =270
                            BorderColor =967423
                            Name ="Label358"
                            Caption ="Short Name"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =345
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =615
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
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11505
                    Top =345
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =967423
                    Name ="clientType"
                    ControlSource ="clientType"
                    RowSourceType ="Value List"
                    RowSource ="TARGET;2ND TARGET;3RD TARGET;DEAD;CLIENT;GUARANTOR"
                    StatusBarText ="Required - Type of Client, i.e.:'Target' or 'Client'"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Client Type is Required"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Required - Type of Client, i.e.:'Target' or 'Client'"
                    GroupTable =2

                    LayoutCachedLeft =11505
                    LayoutCachedTop =345
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =615
                    ColumnStart =7
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =10635
                            Top =345
                            Width =810
                            Height =270
                            BorderColor =967423
                            Name ="Label37"
                            Caption ="Type:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =10635
                            LayoutCachedTop =345
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =615
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =675
                    Width =3675
                    FontSize =9
                    TabIndex =4
                    Name ="clientCoNm2"
                    ControlSource ="clientCoNm2"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =4200
                    LayoutCachedTop =675
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2925
                            Top =675
                            Width =1215
                            Height =240
                            Name ="Label458"
                            Caption ="Co-Lessee:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =2925
                            LayoutCachedTop =675
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =915
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1800
                    Left =11505
                    Top =675
                    Width =1365
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =967423
                    Name ="LOinitial"
                    ControlSource ="LOinitial"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm, Salesperson.SalespersonID FROM Emps INNER JOIN S"
                        "alesperson ON Emps.EmpInt = Salesperson.SPInitials WHERE (((Emps.OBSOLETE)=False"
                        ") AND ((Emps.[LO Exec])=True)) ORDER BY Emps.EmpInt;"
                    ColumnWidths ="360;1440;0"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"HS\""
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    GroupTable =2

                    LayoutCachedLeft =11505
                    LayoutCachedTop =675
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =10635
                            Top =675
                            Width =810
                            Height =240
                            BorderColor =967423
                            Name ="Label585"
                            Caption ="BDEX:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =10635
                            LayoutCachedTop =675
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =915
                            RowStart =1
                            RowEnd =1
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
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =8940
                    Top =675
                    Width =1635
                    FontSize =9
                    TabIndex =5
                    Name ="ClientCONo"
                    ControlSource ="ClientCONo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CoAltDesignation.AltCOid, CoAltDesignation.[CO number], CoAltDesignation."
                        "[Client Grp] FROM CoAltDesignation WHERE (((CoAltDesignation.[Client Grp])=[Form"
                        "s]![ClientsView]![CLIENT GROUP]));"
                    ColumnWidths ="0;7200;360"
                    FontName ="Calibri"
                    GroupTable =2

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =8940
                    LayoutCachedTop =675
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =675
                            Width =950
                            Height =240
                            Name ="Label645"
                            Caption ="Co. Number:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =675
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =915
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =12450
                    Top =1275
                    Width =420
                    TabIndex =11
                    Name ="ONPortal"
                    ControlSource ="ONPortal"
                    StatusBarText ="up load to Portal"
                    GroupTable =2

                    LayoutCachedLeft =12450
                    LayoutCachedTop =1275
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11505
                            Top =1275
                            Width =885
                            Height =240
                            Name ="Label647"
                            Caption ="On Portal:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =11505
                            LayoutCachedTop =1275
                            LayoutCachedWidth =12390
                            LayoutCachedHeight =1515
                            RowStart =3
                            RowEnd =3
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =660
                    Width =2715
                    Height =11100
                    FontSize =8
                    TabIndex =25
                    BoundColumn =1
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Client, clientID FROM ClientsList ORDER BY Client;"
                    ColumnWidths ="2016;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =11760
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =300
                    Width =2715
                    Height =270
                    FontSize =9
                    BorderColor =967423
                    Name ="SearchClient"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =570
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =60
                            Top =75
                            Width =2715
                            Height =240
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label668"
                            Caption ="Type Search HERE:"
                            FontName ="Calibri"
                            LayoutCachedLeft =60
                            LayoutCachedTop =75
                            LayoutCachedWidth =2775
                            LayoutCachedHeight =315
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin EmptyCell
                    Left =10635
                    Top =1275
                    Width =810
                    Name ="EmptyCell910"
                    GroupTable =2
                    LayoutCachedLeft =10635
                    LayoutCachedTop =1275
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =11805
                    Top =105
                    Width =690
                    Height =195
                    FontWeight =700
                    Name ="LockUnLock"
                    Caption ="Locked"
                    FontName ="Calibri"
                    LayoutCachedLeft =11805
                    LayoutCachedTop =105
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =300
                    ThemeFontIndex =1
                    ForeThemeColorIndex =4
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =7050
                    Top =135
                    Width =4755
                    Height =240
                    Name ="Label1802"
                    Caption ="┌─────────────────────┬────────────┬──"
                    LayoutCachedLeft =7050
                    LayoutCachedTop =135
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =4
                End
                Begin EmptyCell
                    Left =10635
                    Top =975
                    Width =810
                    Name ="EmptyCell1887"
                    GroupTable =2
                    LayoutCachedLeft =10635
                    LayoutCachedTop =975
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =1215
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =11505
                    Top =975
                    Width =1365
                    Name ="EmptyCell1888"
                    GroupTable =2
                    LayoutCachedLeft =11505
                    LayoutCachedTop =975
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =1215
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin ComboBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1800
                    Left =14235
                    Top =1080
                    Width =1905
                    Height =270
                    FontSize =9
                    TabIndex =27
                    Name ="OrigBDEX"
                    ControlSource ="OrigBDEX"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm, Salesperson.SalespersonID FROM Emps INNER JOIN S"
                        "alesperson ON Emps.EmpInt = Salesperson.SPInitials WHERE (((Emps.OBSOLETE)=False"
                        ") AND ((Emps.[LO Exec])=True)) ORDER BY Emps.EmpInt;"
                    ColumnWidths ="360;1440;0"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    FontName ="Calibri"
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"

                    LayoutCachedLeft =14235
                    LayoutCachedTop =1080
                    LayoutCachedWidth =16140
                    LayoutCachedHeight =1350
                    ColumnStart =2
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =13020
                            Top =1080
                            Width =1140
                            Height =270
                            Name ="Label1875"
                            Caption ="Orig BDEX:"
                            FontName ="Calibri"
                            LayoutCachedLeft =13020
                            LayoutCachedTop =1080
                            LayoutCachedWidth =14160
                            LayoutCachedHeight =1350
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12960
                    Top =225
                    Width =3240
                    Height =1575
                    Name ="Box1896"
                    LayoutCachedLeft =12960
                    LayoutCachedTop =225
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =1800
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12960
                    Top =1905
                    Width =3956
                    Height =780
                    Name ="Box3"
                    LayoutCachedLeft =12960
                    LayoutCachedTop =1905
                    LayoutCachedWidth =16916
                    LayoutCachedHeight =2685
                    BorderThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =19095
                    Top =2025
                    Width =1759
                    Height =720
                    FontSize =9
                    TabIndex =29
                    Name ="Client_Lease_Summary"
                    Caption ="All Clients in Group Lease Summary"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Complete List of Units Attached To This Client"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =19095
                    LayoutCachedTop =2025
                    LayoutCachedWidth =20854
                    LayoutCachedHeight =2745
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =9211020
                    HoverColor =4754549
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14220
                    Top =1395
                    Width =1905
                    Height =270
                    TabIndex =28
                    ForeColor =4144959
                    Name ="SFAccountID"
                    ControlSource ="SFAccountID"
                    Format ="Short Date"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    Tag ="Required"

                    LayoutCachedLeft =14220
                    LayoutCachedTop =1395
                    LayoutCachedWidth =16125
                    LayoutCachedHeight =1665
                    ColumnStart =13
                    ColumnEnd =13
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =13020
                            Top =1395
                            Width =1155
                            Height =270
                            Name ="Label2092"
                            Caption ="Sales Force ID"
                            FontName ="Calibri"
                            LayoutCachedLeft =13020
                            LayoutCachedTop =1395
                            LayoutCachedWidth =14175
                            LayoutCachedHeight =1665
                            ColumnStart =12
                            ColumnEnd =12
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =13395
                    Top =2385
                    Width =2089
                    Height =240
                    FontWeight =700
                    TabIndex =31
                    Name ="GroupFleetList"
                    Caption ="Fleet List Clients in Group"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Complete List of Units for all Clients within the Group of Clients"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13395
                    LayoutCachedTop =2385
                    LayoutCachedWidth =15484
                    LayoutCachedHeight =2625
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =9211020
                    HoverColor =4754549
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =13395
                    Top =1965
                    Width =2089
                    Height =240
                    FontWeight =700
                    TabIndex =30
                    Name ="ClientFleetList"
                    Caption ="Fleet List This Client Only"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Complete List of Units Attached To This Client"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13395
                    LayoutCachedTop =1965
                    LayoutCachedWidth =15484
                    LayoutCachedHeight =2205
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =9211020
                    HoverColor =4754549
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =13050
                    Top =120
                    Width =1425
                    Height =240
                    FontWeight =700
                    Name ="Label2141"
                    Caption ="BDEX Details:"
                    FontName ="Calibri"
                    LayoutCachedLeft =13050
                    LayoutCachedTop =120
                    LayoutCachedWidth =14475
                    LayoutCachedHeight =360
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8940
                    Top =1875
                    Width =3930
                    FontSize =9
                    TabIndex =15
                    Name ="clientIndustry"
                    ControlSource ="clientIndustry"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Industry"
                    StatusBarText ="link to industry lookup table"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =1875
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =2115
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =1875
                            Width =950
                            Height =240
                            Name ="Label618"
                            Caption ="Industry:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =1875
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =2115
                            RowStart =5
                            RowEnd =5
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
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8940
                    Top =2175
                    Width =3930
                    FontSize =9
                    TabIndex =19
                    Name ="clientAnnualSales"
                    ControlSource ="clientAnnualSales"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Networth if client type is Bank"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =2175
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =2415
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =2175
                            Width =950
                            Height =240
                            Name ="Label617"
                            Caption =" Sales:"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =2175
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =2415
                            RowStart =6
                            RowEnd =6
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
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =8940
                    Top =2475
                    Width =3930
                    FontSize =9
                    TabIndex =22
                    Name ="CPMModel"
                    ControlSource ="CPMModel"
                    RowSourceType ="Table/Query"
                    RowSource ="Select * From CMPVerticals"
                    ColumnWidths ="4320;0"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =8940
                    LayoutCachedTop =2475
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =2715
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =7935
                            Top =2475
                            Width =950
                            Height =240
                            Name ="CMPModel_Label"
                            Caption ="MnR Vertical"
                            FontName ="Calibri"
                            GroupTable =2
                            LayoutCachedLeft =7935
                            LayoutCachedTop =2475
                            LayoutCachedWidth =8885
                            LayoutCachedHeight =2715
                            RowStart =7
                            RowEnd =7
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =7
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =6030
                    Top =2775
                    Width =705
                    Height =255
                    Name ="EmptyCell2233"
                    GroupTable =2
                    LayoutCachedLeft =6030
                    LayoutCachedTop =2775
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =3030
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6795
                    Top =2775
                    Width =1080
                    Height =255
                    Name ="EmptyCell2234"
                    GroupTable =2
                    LayoutCachedLeft =6795
                    LayoutCachedTop =2775
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =3030
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7935
                    Top =2775
                    Width =950
                    Height =255
                    Name ="EmptyCell2270"
                    GroupTable =2
                    LayoutCachedLeft =7935
                    LayoutCachedTop =2775
                    LayoutCachedWidth =8885
                    LayoutCachedHeight =3030
                    RowStart =8
                    RowEnd =8
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8940
                    Top =2775
                    Width =3930
                    Height =255
                    Name ="EmptyCell2271"
                    GroupTable =2
                    LayoutCachedLeft =8940
                    LayoutCachedTop =2775
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =3030
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =15735
                    Top =2145
                    Width =2089
                    Height =240
                    FontWeight =700
                    TabIndex =33
                    Name ="FleetListProfile"
                    Caption ="Fleet List Profile"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Complete List of Units for all Clients within the Group of Clients"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15735
                    LayoutCachedTop =2145
                    LayoutCachedWidth =17824
                    LayoutCachedHeight =2385
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =9211020
                    HoverColor =4754549
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17940
                    Top =960
                    Width =1020
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =34
                    Name ="ExtendedRentdiscount"
                    ControlSource ="ClientComm%"
                    Format ="Percent"
                    DefaultValue ="0.8"
                    FontName ="Segoe UI"
                    Tag ="Acct"

                    LayoutCachedLeft =17940
                    LayoutCachedTop =960
                    LayoutCachedWidth =18960
                    LayoutCachedHeight =1320
                    RowStart =10
                    RowEnd =10
                    ColumnStart =7
                    ColumnEnd =7
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =17940
                            Top =540
                            Width =1005
                            Height =420
                            Name ="Label3617"
                            Caption ="Extended Rent Rate"
                            FontName ="Segoe UI"
                            Tag ="Acct"
                            LayoutCachedLeft =17940
                            LayoutCachedTop =540
                            LayoutCachedWidth =18945
                            LayoutCachedHeight =960
                            ForeThemeColorIndex =2
                        End
                    End
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
' See "ClientsView.cls"
