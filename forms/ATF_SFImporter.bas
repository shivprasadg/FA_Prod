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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10365
    DatasheetFontHeight =11
    ItemSuffix =1334
    Left =4845
    Top =2715
    Right =15465
    Bottom =12375
    RecSrcDt = Begin
        0xf2a31b5e1544e640
    End
    RecordSource ="SELECT DISTINCT vw_SFKeys.AccountId, vw_SFKeys.OpportunityId, vw_SFKeys.LPAFId, "
        "vw_SFKeys.SpecId, vw_SFKeys.ClientGroupId, vw_SFKeys.clientID, vw_SFKeys.Client,"
        " vw_SFKeys.ClientSN, vw_SFKeys.BDEX, vw_SFKeys.Dealer, vw_SFKeys.FAID, vw_SFKeys"
        ".FAIDAC, vw_SFKeys.LPAFStart, vw_SFKeys.LPAFDate, vw_SFKeys.NDA_Signed, vw_SFKey"
        "s.OutcomeDate, vw_SFKeys.TaxID, vw_SFKeys.UnitCount, vw_SFKeys.VendorID, vw_SFKe"
        "ys.VendorSN, vw_SFI_SchDealSpec.LeaseDescp, vw_SFI_SchDealSpec.SchShortDesc, vw_"
        "SFI_SchDealSpec.SpecGroups, vw_SFI_SchDealSpec.AvgUnitCost, vw_SFI_SchDealSpec.D"
        "ealAmount, Salesperson.SalespersonID, vw_SFKeys.LeaseType FROM (vw_SFKeys INNER "
        "JOIN vw_SFI_SchDealSpec ON (vw_SFKeys.FAID = vw_SFI_SchDealSpec.FAID) AND (vw_SF"
        "Keys.OpportunityId = vw_SFI_SchDealSpec.OpportunityId)) INNER JOIN Salesperson O"
        "N vw_SFI_SchDealSpec.BDEX = Salesperson.SPInitials WHERE (((vw_SFKeys.FAIDAC) Is"
        " Null));"
    Caption ="Import Saleforce.com Awarded Deal to ATLAAS"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
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
        Begin FormHeader
            Height =675
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =-30
                    Top =270
                    Width =5040
                    Height =360
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    ForeColor =0
                    Name ="SF_ClientLegalName"
                    ControlSource ="Client"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =-30
                    LayoutCachedTop =270
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =630
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =-30
                    Width =5040
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label1066"
                    Caption ="Importing From"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =-30
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =270
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5220
                    Width =5055
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label979"
                    Caption ="To "
                    FontName ="Segoe UI"
                    LayoutCachedLeft =5220
                    LayoutCachedWidth =10275
                    LayoutCachedHeight =270
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Top =270
                    Width =5055
                    Height =360
                    ColumnOrder =1
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =14211288
                    ForeColor =0
                    Name ="ATLAASClient"
                    ControlSource ="=DLookUp(\"clientCompanyName\",\"Clients\",\"ClientID=\" & [LesseeID])"
                    FontName ="Segoe UI"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000000000003000000000000001500000001010000 ,
                        0x00000000ffc20e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530046005f0043006c00690065006e0074004c006500670061006c004e00 ,
                        0x61006d0065005d0000000000
                    End

                    LayoutCachedLeft =5220
                    LayoutCachedTop =270
                    LayoutCachedWidth =10275
                    LayoutCachedHeight =630
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000101000000000000ffc20e00140000005b00 ,
                        0x530046005f0043006c00690065006e0074004c006500670061006c004e006100 ,
                        0x6d0065005d00000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8160
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Tab
                    OverlapFlags =85
                    Style =2
                    Left =12
                    Top =456
                    Width =10344
                    Height =7692
                    FontWeight =700
                    Name ="TabCtrl1"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12
                    LayoutCachedTop =456
                    LayoutCachedWidth =10356
                    LayoutCachedHeight =8148
                    ThemeFontIndex =-1
                    UseTheme =255
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =12566463
                    HoverForeColor =8872517
                    PressedForeColor =0
                    ForeColor =0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =75
                            Top =510
                            Width =10215
                            Height =7575
                            Name ="Step1"
                            Caption ="Step 1"
                            LayoutCachedLeft =75
                            LayoutCachedTop =510
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =8085
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =10
                                    Left =120
                                    Top =4455
                                    Width =9975
                                    Height =3180
                                    FontSize =10
                                    ForeColor =0
                                    BorderColor =8355711
                                    Name ="FAIDLPAFList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT vw_SFKeys.FAID, vw_SFKeys.Client AS [Account Name], vw_SFKeys.LP"
                                        "AFDate AS [Actual Close], 1 AS Sch, vw_SFI_SchDealSpec.SpecGroups AS Groups, vw_"
                                        "SFI_SchDealSpec.UnitCount AS Units, vw_SFI_SchDealSpec.SchShortDesc AS [Opportun"
                                        "ity Name], vw_SFKeys.OpportunityId, vw_SFKeys.ClientGroupId, vw_SFKeys.clientID,"
                                        " vw_SFKeys.AccountId, \"N/A\" AS SpecId, vw_SFI_SchDealSpec.FAIDAC FROM vw_SFKey"
                                        "s INNER JOIN vw_SFI_SchDealSpec ON (vw_SFKeys.OpportunityId = vw_SFI_SchDealSpec"
                                        ".OpportunityId) AND (vw_SFKeys.FAID = vw_SFI_SchDealSpec.FAID) WHERE (((vw_SFI_S"
                                        "chDealSpec.FAIDAC)=\"import\")) ORDER BY vw_SFKeys.Client, vw_SFKeys.LPAFDate DE"
                                        "SC;"
                                    ColumnWidths ="1080;3456;1440;576;576;576;2160;0;0;0"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =4455
                                    LayoutCachedWidth =10095
                                    LayoutCachedHeight =7635
                                    BackThemeColorIndex =-1
                                    BorderShade =50.0
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =420
                                    Top =2175
                                    Width =5430
                                    Height =285
                                    FontWeight =700
                                    TabIndex =1
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text1264"
                                    ControlSource ="Dealer"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =2175
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =2460
                                    RowStart =7
                                    RowEnd =7
                                    ColumnEnd =6
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =420
                                    Top =990
                                    Width =2145
                                    Height =285
                                    FontWeight =700
                                    TabIndex =2
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="clientCoNm2"
                                    ControlSource ="FAID"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =990
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =1275
                                    RowStart =3
                                    RowEnd =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5925
                                    Top =1575
                                    Width =1905
                                    Height =285
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    ForeColor =0
                                    Name ="ActualCloseDate"
                                    ControlSource ="OutcomeDate"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    ControlTipText ="In SFDC, field is Actual Close Date on the Opportunity form"

                                    LayoutCachedLeft =5925
                                    LayoutCachedTop =1575
                                    LayoutCachedWidth =7830
                                    LayoutCachedHeight =1860
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =2
                                    BorderShade =50.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2625
                                    Top =990
                                    Width =855
                                    Height =285
                                    FontWeight =700
                                    TabIndex =4
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text993"
                                    ControlSource ="BDEX"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =2625
                                    LayoutCachedTop =990
                                    LayoutCachedWidth =3480
                                    LayoutCachedHeight =1275
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =420
                                    Top =1575
                                    Width =5430
                                    Height =285
                                    FontWeight =700
                                    TabIndex =5
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text1054"
                                    ControlSource ="SchShortDesc"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =1575
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =1860
                                    RowStart =5
                                    RowEnd =5
                                    ColumnEnd =6
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2625
                                    Top =2775
                                    Width =5205
                                    Height =285
                                    FontWeight =700
                                    TabIndex =6
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text1016"
                                    ControlSource ="LeaseDescp"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =2625
                                    LayoutCachedTop =2775
                                    LayoutCachedWidth =7830
                                    LayoutCachedHeight =3060
                                    RowStart =7
                                    RowEnd =7
                                    ColumnEnd =6
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =420
                                    Top =2775
                                    Width =2145
                                    Height =285
                                    FontWeight =700
                                    TabIndex =7
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text1270"
                                    ControlSource ="DealAmount"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =2775
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =3060
                                    RowStart =7
                                    RowEnd =7
                                    ColumnEnd =6
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    CurrencySymbol ="$"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =420
                                    Top =1935
                                    Width =5430
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1263"
                                    Caption ="Vendor"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =1935
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =2175
                                    RowStart =6
                                    RowEnd =6
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =120
                                    Top =4080
                                    Width =3315
                                    Height =435
                                    FontSize =16
                                    FontWeight =700
                                    BorderColor =16777215
                                    ForeColor =0
                                    Name ="Label12"
                                    Caption ="Select an FAID to Import"
                                    LayoutCachedLeft =120
                                    LayoutCachedTop =4080
                                    LayoutCachedWidth =3435
                                    LayoutCachedHeight =4515
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =420
                                    Top =750
                                    Width =2145
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label458"
                                    Caption ="FAID"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =750
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =990
                                    RowStart =2
                                    RowEnd =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =5925
                                    Top =1335
                                    Width =1905
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1015"
                                    Caption ="Actual Close Date"
                                    LayoutCachedLeft =5925
                                    LayoutCachedTop =1335
                                    LayoutCachedWidth =7830
                                    LayoutCachedHeight =1575
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =2625
                                    Top =750
                                    Width =855
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label992"
                                    Caption ="BDEX"
                                    LayoutCachedLeft =2625
                                    LayoutCachedTop =750
                                    LayoutCachedWidth =3480
                                    LayoutCachedHeight =990
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =420
                                    Top =1335
                                    Width =5430
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1053"
                                    Caption ="Opportunity Name"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =1335
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =1575
                                    RowStart =4
                                    RowEnd =4
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =2625
                                    Top =2535
                                    Width =5205
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1028"
                                    Caption ="Lease Type"
                                    LayoutCachedLeft =2625
                                    LayoutCachedTop =2535
                                    LayoutCachedWidth =7830
                                    LayoutCachedHeight =2775
                                    RowStart =6
                                    RowEnd =6
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =420
                                    Top =2535
                                    Width =2145
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1269"
                                    Caption ="Deal Amount"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =2535
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =2775
                                    RowStart =6
                                    RowEnd =6
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    Left =255
                                    Top =705
                                    Width =8385
                                    Height =2490
                                    Name ="whitebox"
                                    LayoutCachedLeft =255
                                    LayoutCachedTop =705
                                    LayoutCachedWidth =8640
                                    LayoutCachedHeight =3195
                                    BackThemeColorIndex =-1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =510
                            Width =10215
                            Height =7575
                            Name ="Step2"
                            Caption ="Step 2"
                            LayoutCachedLeft =75
                            LayoutCachedTop =510
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =8085
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =119
                                    Top =660
                                    Width =8265
                                    Height =435
                                    FontSize =14
                                    FontWeight =700
                                    BorderColor =16777215
                                    ForeColor =0
                                    Name ="Label314"
                                    Caption ="Select An Existing Master Lease or Add a New Master Lease"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =119
                                    LayoutCachedTop =660
                                    LayoutCachedWidth =8384
                                    LayoutCachedHeight =1095
                                    RowStart =4
                                    RowEnd =4
                                    ThemeFontIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =600
                                    Top =1092
                                    Width =1530
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label486"
                                    Caption ="FAID"
                                    GroupTable =3
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =2130
                                    LayoutCachedHeight =1332
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =2196
                                    Top =1092
                                    Width =5820
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label487"
                                    Caption ="Short Description"
                                    GroupTable =3
                                    LayoutCachedLeft =2196
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =1332
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =600
                                    Top =1392
                                    Width =1530
                                    Height =360
                                    FontWeight =700
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text488"
                                    ControlSource ="FAID"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =3

                                    LayoutCachedLeft =600
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =2130
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =1
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2196
                                    Top =1392
                                    Width =5820
                                    Height =360
                                    FontWeight =700
                                    TabIndex =1
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text489"
                                    ControlSource ="SchShortDesc"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                                    GroupTable =3

                                    LayoutCachedLeft =2196
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =600
                                    Top =1812
                                    Width =7416
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label490"
                                    Caption ="Lessee"
                                    GroupTable =3
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =600
                                    Top =2112
                                    Width =7416
                                    Height =360
                                    FontWeight =700
                                    TabIndex =2
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text491"
                                    ControlSource ="=DLookUp(\"clientCompanyName\",\"Clients\",\"ClientID=\" & [LesseeID])"
                                    StatusBarText ="Primary Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Primary Lessee"
                                    GroupTable =3

                                    LayoutCachedLeft =600
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =600
                                    Top =2532
                                    Width =7416
                                    Height =240
                                    FontSize =10
                                    Name ="Label504"
                                    Caption ="Select the Master Lease to use for the FAID below"
                                    GroupTable =3
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =2772
                                    RowStart =4
                                    RowEnd =4
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3960
                                    Left =600
                                    Top =2832
                                    Width =7416
                                    Height =396
                                    FontWeight =700
                                    TabIndex =3
                                    BorderColor =14277081
                                    ForeColor =6108695
                                    Name ="MLID"
                                    RowSourceType ="Value List"
                                    RowSource ="0;'Create A Master Lease On Import';0;290;'1';4400"
                                    ColumnWidths ="0;3960;0"
                                    StatusBarText ="Master Lease"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="MLID is required"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Required - Master Lease"
                                    GroupTable =3
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =600
                                    LayoutCachedTop =2832
                                    LayoutCachedWidth =8016
                                    LayoutCachedHeight =3228
                                    RowStart =5
                                    RowEnd =5
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    ThemeFontIndex =-1
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeShade =100.0
                                    GroupTable =3
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =8025
                                    Top =2880
                                    Width =300
                                    Height =285
                                    FontSize =12
                                    Name ="Label571"
                                    Caption ="◄"
                                    LayoutCachedLeft =8025
                                    LayoutCachedTop =2880
                                    LayoutCachedWidth =8325
                                    LayoutCachedHeight =3165
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =510
                            Width =10215
                            Height =7580
                            Name ="Step3"
                            Caption ="Step 3"
                            LayoutCachedLeft =75
                            LayoutCachedTop =510
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =8090
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =119
                                    Top =660
                                    Width =5700
                                    Height =435
                                    FontSize =16
                                    FontWeight =700
                                    BorderColor =16777215
                                    ForeColor =0
                                    Name ="Label402"
                                    Caption ="Add Needed Information"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =119
                                    LayoutCachedTop =660
                                    LayoutCachedWidth =5819
                                    LayoutCachedHeight =1095
                                    ColumnEnd =1
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =1092
                                    Width =1536
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label318"
                                    Caption ="FAID"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =1332
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =2172
                                    Top =1092
                                    Width =5484
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label405"
                                    Caption ="Short Description"
                                    GroupTable =12
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =1332
                                    ColumnStart =1
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =1092
                                    Width =1536
                                    Name ="EmptyCell1329"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =1332
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =1092
                                    Width =276
                                    Name ="EmptyCell755"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =1332
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =1392
                                    Width =1536
                                    Height =360
                                    FontWeight =700
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="FAID"
                                    ControlSource ="FAID"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2172
                                    Top =1392
                                    Width =5484
                                    Height =360
                                    FontWeight =700
                                    TabIndex =1
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="ATFDesc"
                                    ControlSource ="SchShortDesc"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                                    GroupTable =12

                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =1392
                                    Width =1536
                                    Height =360
                                    Name ="EmptyCell1328"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =1392
                                    Width =276
                                    Height =360
                                    Name ="EmptyCell756"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =1812
                                    Width =7080
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label304"
                                    Caption ="Lessee"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =7716
                                    Top =1812
                                    Width =1536
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1305"
                                    Caption ="LeaseType"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =1812
                                    Width =276
                                    Name ="EmptyCell757"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =2112
                                    Width =7080
                                    Height =360
                                    FontWeight =700
                                    TabIndex =2
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Lessee"
                                    ControlSource ="=DLookUp(\"clientCompanyName\",\"Clients\",\"ClientID=\" & [LesseeID])"
                                    StatusBarText ="Primary Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Primary Lessee"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7716
                                    Top =2112
                                    Width =1536
                                    Height =360
                                    FontWeight =700
                                    TabIndex =3
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text1306"
                                    ControlSource ="LeaseType"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =12

                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =2112
                                    Width =276
                                    Height =360
                                    Name ="EmptyCell758"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =2532
                                    Width =7080
                                    Height =204
                                    Name ="EmptyCell946"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =2736
                                    RowStart =4
                                    RowEnd =4
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =2532
                                    Width =1536
                                    Height =204
                                    Name ="EmptyCell1313"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =2736
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =2532
                                    Width =276
                                    Height =204
                                    Name ="EmptyCell951"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =2736
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =2796
                                    Width =7080
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label485"
                                    Caption ="Master Lease"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2796
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =3036
                                    RowStart =5
                                    RowEnd =5
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =2796
                                    Width =1536
                                    Name ="EmptyCell1314"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =2796
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =3036
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =2796
                                    Width =276
                                    Name ="EmptyCell759"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =2796
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =3036
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =3096
                                    Width =7080
                                    Height =396
                                    FontWeight =700
                                    TabIndex =4
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="MLID2"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="MLID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3096
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =3492
                                    RowStart =6
                                    RowEnd =6
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =3096
                                    Width =1536
                                    Height =396
                                    Name ="EmptyCell1315"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =3096
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =3492
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =3096
                                    Width =276
                                    Height =396
                                    Name ="EmptyCell760"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =3096
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =3492
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =3552
                                    Width =7080
                                    Height =192
                                    Name ="EmptyCell934"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3552
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =3744
                                    RowStart =7
                                    RowEnd =7
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =3552
                                    Width =1536
                                    Height =192
                                    Name ="EmptyCell1316"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =3552
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =3744
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =3552
                                    Width =276
                                    Height =192
                                    Name ="EmptyCell939"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =3552
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =3744
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =3804
                                    Width =7080
                                    Height =240
                                    FontSize =10
                                    Name ="Label305"
                                    Caption ="Co-Lessee"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =4044
                                    RowStart =8
                                    RowEnd =8
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =3804
                                    Width =1536
                                    Name ="EmptyCell1317"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =4044
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =3804
                                    Width =276
                                    Name ="EmptyCell761"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =4044
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =576
                                    Top =4104
                                    Width =7080
                                    Height =360
                                    FontWeight =700
                                    TabIndex =5
                                    BorderColor =14277081
                                    ForeColor =6108695
                                    Name ="CoLesseeID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Clients.clientID, Clients.clientCompanyName AS CoLessee, Clients.ClientSh"
                                        "Nm AS ShortName, Clients.ClientGroupId FROM Clients WHERE (((Clients.ClientGroup"
                                        "Id)=DLookUp(\"ClientGroupId\",\"Clients\",\"ClientID=\" & [LesseeID]))) ORDER BY"
                                        " Clients.clientCompanyName;"
                                    ColumnWidths ="0;3168;1152"
                                    StatusBarText ="Co-Lessee"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4104
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =4464
                                    RowStart =9
                                    RowEnd =9
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    ThemeFontIndex =-1
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeShade =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =4104
                                    Width =1536
                                    Height =360
                                    Name ="EmptyCell1318"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =4104
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =4464
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =4104
                                    Width =276
                                    Height =360
                                    Name ="EmptyCell762"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =4104
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =4464
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =4524
                                    Width =7080
                                    Height =192
                                    Name ="EmptyCell928"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4524
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =4716
                                    RowStart =10
                                    RowEnd =10
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =4524
                                    Width =1536
                                    Height =192
                                    Name ="EmptyCell1319"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =4524
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =4716
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =4524
                                    Width =276
                                    Height =192
                                    Name ="EmptyCell933"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =4524
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =4716
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =4776
                                    Width =7080
                                    Height =240
                                    FontSize =10
                                    Name ="Label774"
                                    Caption ="Sub-Lessee"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4776
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =5016
                                    RowStart =11
                                    RowEnd =11
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =4776
                                    Width =1536
                                    Name ="EmptyCell1320"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =4776
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =5016
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =4776
                                    Width =276
                                    Name ="EmptyCell834"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =4776
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =5016
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =5076
                                    Width =7080
                                    Height =360
                                    FontWeight =700
                                    TabIndex =6
                                    BorderColor =14277081
                                    ForeColor =6108695
                                    Name ="SubLessee1"
                                    StatusBarText ="Co-Lessee"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5076
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =5436
                                    RowStart =12
                                    RowEnd =12
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =5076
                                    Width =1536
                                    Height =360
                                    Name ="EmptyCell1321"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =5076
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =5436
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =5076
                                    Width =276
                                    Height =360
                                    Name ="EmptyCell838"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =5076
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =5436
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =5496
                                    Width =7080
                                    Height =192
                                    Name ="EmptyCell922"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5496
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =5688
                                    RowStart =13
                                    RowEnd =13
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =5496
                                    Width =1536
                                    Height =192
                                    Name ="EmptyCell1322"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =5496
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =5688
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =5496
                                    Width =276
                                    Height =192
                                    Name ="EmptyCell927"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =5496
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =5688
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =5748
                                    Width =7080
                                    Height =240
                                    FontSize =10
                                    Name ="Label507"
                                    Caption ="Guarantor"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5748
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =5988
                                    RowStart =14
                                    RowEnd =14
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =5748
                                    Width =1536
                                    Name ="EmptyCell1323"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =5748
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =5988
                                    RowStart =14
                                    RowEnd =14
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =5748
                                    Width =276
                                    Name ="EmptyCell763"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =5748
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =5988
                                    RowStart =14
                                    RowEnd =14
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =576
                                    Top =6048
                                    Width =7080
                                    Height =360
                                    FontWeight =700
                                    TabIndex =7
                                    BorderColor =14277081
                                    ForeColor =6108695
                                    Name ="GuarantorID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Clients.clientID, Clients.clientCompanyName AS CoLessee, Clients.ClientSh"
                                        "Nm AS ShortName, Clients.ClientGroupId FROM Clients ORDER BY Clients.clientCompa"
                                        "nyName;"
                                    ColumnWidths ="0;3168;1152"
                                    StatusBarText ="Co-Lessee"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6048
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =6408
                                    RowStart =15
                                    RowEnd =15
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    ThemeFontIndex =-1
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeShade =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =6048
                                    Width =1536
                                    Height =360
                                    Name ="EmptyCell1324"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =6048
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =6408
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =6048
                                    Width =276
                                    Height =360
                                    Name ="EmptyCell764"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =6048
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =6408
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =6468
                                    Width =1536
                                    Height =192
                                    Name ="EmptyCell940"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =2172
                                    Top =6468
                                    Width =1452
                                    Height =192
                                    Name ="EmptyCell941"
                                    GroupTable =12
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =3696
                                    Top =6468
                                    Width =924
                                    Height =192
                                    Name ="EmptyCell942"
                                    GroupTable =12
                                    LayoutCachedLeft =3696
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =4620
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =4680
                                    Top =6468
                                    Width =1452
                                    Height =192
                                    Name ="EmptyCell943"
                                    GroupTable =12
                                    LayoutCachedLeft =4680
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =6132
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =6204
                                    Top =6468
                                    Width =1452
                                    Height =192
                                    Name ="EmptyCell944"
                                    GroupTable =12
                                    LayoutCachedLeft =6204
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =6468
                                    Width =1536
                                    Height =192
                                    Name ="EmptyCell1325"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =6468
                                    Width =276
                                    Height =192
                                    Name ="EmptyCell945"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =6660
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =576
                                    Top =6720
                                    Width =1536
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label97"
                                    Caption ="LPAF Date"
                                    GroupTable =12
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2172
                                    Top =6720
                                    Width =1452
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label585"
                                    Caption ="BDEX"
                                    GroupTable =12
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =3696
                                    Top =6720
                                    Width =924
                                    Name ="EmptyCell631"
                                    GroupTable =12
                                    LayoutCachedLeft =3696
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =4620
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4680
                                    Top =6720
                                    Width =1452
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label384"
                                    Caption ="ADMIN"
                                    GroupTable =12
                                    LayoutCachedLeft =4680
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =6132
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =6204
                                    Top =6720
                                    Width =1452
                                    Height =240
                                    FontSize =10
                                    Name ="Label96"
                                    Caption ="ATF Date"
                                    GroupTable =12
                                    LayoutCachedLeft =6204
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =6720
                                    Width =1536
                                    Name ="EmptyCell1326"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =6720
                                    Width =276
                                    Name ="EmptyCell765"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =6720
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =6960
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =7032
                                    Width =1536
                                    Height =336
                                    FontWeight =700
                                    TabIndex =8
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="LPAFDate"
                                    ControlSource ="=CDate([OutcomeDate])"
                                    Format ="Short Date"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    ControlTipText ="In SFDC, field is Actual Close Date on the Opportunity form"
                                    GroupTable =12

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2172
                                    Top =7032
                                    Width =1452
                                    Height =336
                                    FontWeight =700
                                    TabIndex =9
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="BDEX2"
                                    ControlSource ="=DLookUp(\"SPInitials\",\"SalesPerson\",\"SalesPersonID=\" & [SalesPersonID])"
                                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                                    GroupTable =12

                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =3696
                                    Top =7032
                                    Width =924
                                    Height =336
                                    TabIndex =10
                                    Name ="BDEX"
                                    ControlSource ="SalesPersonID"
                                    GroupTable =12

                                    LayoutCachedLeft =3696
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =4620
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4680
                                    Top =7032
                                    Width =1452
                                    Height =336
                                    FontWeight =700
                                    TabIndex =11
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="ATFAdmin"
                                    StatusBarText ="Required - ATF Admin"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="Admin is Required"
                                    DefaultValue ="=[FAUserInit]"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Required - ATF Admin"
                                    GroupTable =12

                                    LayoutCachedLeft =4680
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =6132
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =2
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6204
                                    Top =7032
                                    Width =1452
                                    Height =336
                                    FontWeight =700
                                    TabIndex =12
                                    BackColor =16249583
                                    BorderColor =6108695
                                    ForeColor =0
                                    Name ="ATFDate"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Required - ATF Date"
                                    GroupTable =12

                                    LayoutCachedLeft =6204
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =7656
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =2
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =7716
                                    Top =7032
                                    Width =1536
                                    Height =336
                                    Name ="EmptyCell1327"
                                    GroupTable =12
                                    LayoutCachedLeft =7716
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =9252
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                                Begin EmptyCell
                                    Left =9312
                                    Top =7032
                                    Width =276
                                    Height =336
                                    Name ="EmptyCell952"
                                    GroupTable =12
                                    LayoutCachedLeft =9312
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =9588
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =2
                                    GroupTable =12
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =510
                            Width =10215
                            Height =7575
                            Name ="Step4"
                            Caption ="Step 4"
                            LayoutCachedLeft =75
                            LayoutCachedTop =510
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =8085
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =119
                                    Top =660
                                    Width =5700
                                    Height =435
                                    FontSize =16
                                    FontWeight =700
                                    BorderColor =16777215
                                    ForeColor =0
                                    Name ="Label534"
                                    Caption ="Review and Import "
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =119
                                    LayoutCachedTop =660
                                    LayoutCachedWidth =5819
                                    LayoutCachedHeight =1095
                                    ColumnEnd =1
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =1092
                                    Width =1536
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label644"
                                    Caption ="FAID"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =1332
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =2172
                                    Top =1092
                                    Width =5340
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label645"
                                    Caption ="Short Description"
                                    GroupTable =32
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =1332
                                    ColumnStart =1
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =7572
                                    Top =1092
                                    Width =2292
                                    Height =240
                                    FontSize =10
                                    Name ="Label1303"
                                    Caption ="Enter NEW Schedule Num"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =1092
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =1332
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    BackThemeColorIndex =2
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =1392
                                    Width =1536
                                    Height =360
                                    FontWeight =700
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text646"
                                    ControlSource ="FAID"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2172
                                    Top =1392
                                    Width =5340
                                    Height =360
                                    FontWeight =700
                                    TabIndex =1
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text647"
                                    ControlSource ="SchShortDesc"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                                    GroupTable =32

                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =9942
                                            Top =1452
                                            Width =300
                                            Height =360
                                            FontSize =12
                                            Name ="Label1281"
                                            Caption ="◄"
                                            LayoutCachedLeft =9942
                                            LayoutCachedTop =1452
                                            LayoutCachedWidth =10242
                                            LayoutCachedHeight =1812
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            BorderShade =85.0
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7572
                                    Top =1392
                                    Width =2292
                                    Height =360
                                    FontWeight =700
                                    TabIndex =2
                                    BorderColor =14277081
                                    Name ="NewScheduleName"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =1392
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =1752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =1812
                                    Width =6936
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label648"
                                    Caption ="Lessee"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =7572
                                    Top =1812
                                    Width =2292
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label1332"
                                    Caption ="LeaseType"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =1812
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =2052
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =2112
                                    Width =6936
                                    Height =360
                                    FontWeight =700
                                    TabIndex =3
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="Text649"
                                    ControlSource ="=DLookUp(\"clientCompanyName\",\"Clients\",\"ClientID=\" & [LesseeID])"
                                    StatusBarText ="Primary Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Primary Lessee"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7572
                                    Top =2112
                                    Width =2292
                                    Height =360
                                    FontWeight =700
                                    TabIndex =4
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="LeaseType"
                                    ControlSource ="LeaseType"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="FAID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =2112
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =2472
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =2532
                                    Width =6936
                                    Height =204
                                    Name ="EmptyCell953"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =2736
                                    RowStart =4
                                    RowEnd =4
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =2532
                                    Width =2292
                                    Height =204
                                    Name ="EmptyCell1288"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =2532
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =2736
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =2796
                                    Width =6936
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label650"
                                    Caption ="Master Lease"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =2796
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =3036
                                    RowStart =5
                                    RowEnd =5
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =2796
                                    Width =2292
                                    Name ="EmptyCell1289"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =2796
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =3036
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =3096
                                    Width =6936
                                    Height =396
                                    FontWeight =700
                                    TabIndex =5
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="MLID3"
                                    StatusBarText ="deal id"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="MLID is required"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3096
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =3492
                                    RowStart =6
                                    RowEnd =6
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =3096
                                    Width =2292
                                    Height =396
                                    Name ="EmptyCell1290"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =3096
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =3492
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =3552
                                    Width =6936
                                    Height =192
                                    Name ="EmptyCell958"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3552
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =3744
                                    RowStart =7
                                    RowEnd =7
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =3552
                                    Width =2292
                                    Height =192
                                    Name ="EmptyCell1291"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =3552
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =3744
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =3804
                                    Width =6936
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label652"
                                    Caption ="Co-Lessee"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =4044
                                    RowStart =8
                                    RowEnd =8
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =3804
                                    Width =2292
                                    Name ="EmptyCell1292"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =4044
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =4104
                                    Width =6936
                                    Height =360
                                    FontWeight =700
                                    TabIndex =6
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="CoLesseeID3"
                                    StatusBarText ="Co-Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4104
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =4464
                                    RowStart =9
                                    RowEnd =9
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =4104
                                    Width =2292
                                    Height =360
                                    Name ="EmptyCell1293"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =4104
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =4464
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =4524
                                    Width =6936
                                    Height =192
                                    Name ="EmptyCell963"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4524
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =4716
                                    RowStart =10
                                    RowEnd =10
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =4524
                                    Width =2292
                                    Height =192
                                    Name ="EmptyCell1294"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =4524
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =4716
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =4776
                                    Width =6936
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label842"
                                    Caption ="Sub-Lessee"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =4776
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =5016
                                    RowStart =11
                                    RowEnd =11
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =4776
                                    Width =2292
                                    Name ="EmptyCell1295"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =4776
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =5016
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =5076
                                    Width =6936
                                    Height =360
                                    FontWeight =700
                                    TabIndex =7
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="SubLessee"
                                    StatusBarText ="Co-Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5076
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =5436
                                    RowStart =12
                                    RowEnd =12
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =5076
                                    Width =2292
                                    Height =360
                                    Name ="EmptyCell1296"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =5076
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =5436
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =5496
                                    Width =6936
                                    Height =192
                                    Name ="EmptyCell968"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5496
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =5688
                                    RowStart =13
                                    RowEnd =13
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =5496
                                    Width =2292
                                    Height =192
                                    Name ="EmptyCell1297"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =5496
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =5688
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =576
                                    Top =5748
                                    Width =6936
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label654"
                                    Caption ="Guarantor"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =5748
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =5988
                                    RowStart =14
                                    RowEnd =14
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =5748
                                    Width =2292
                                    Name ="EmptyCell1298"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =5748
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =5988
                                    RowStart =14
                                    RowEnd =14
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =6048
                                    Width =6936
                                    Height =360
                                    FontWeight =700
                                    TabIndex =8
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="GuarantorID3"
                                    StatusBarText ="Co-Lessee"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Co-Lessee"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6048
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =6408
                                    RowStart =15
                                    RowEnd =15
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =6048
                                    Width =2292
                                    Height =360
                                    Name ="EmptyCell1299"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =6048
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =6408
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =576
                                    Top =6468
                                    Width =1536
                                    Height =204
                                    Name ="EmptyCell973"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =2172
                                    Top =6468
                                    Width =1452
                                    Height =204
                                    Name ="EmptyCell974"
                                    GroupTable =32
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =3684
                                    Top =6468
                                    Width =912
                                    Height =204
                                    Name ="EmptyCell975"
                                    GroupTable =32
                                    LayoutCachedLeft =3684
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =4596
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =4656
                                    Top =6468
                                    Width =1428
                                    Height =204
                                    Name ="EmptyCell976"
                                    GroupTable =32
                                    LayoutCachedLeft =4656
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =6084
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =6144
                                    Top =6468
                                    Width =1368
                                    Height =204
                                    Name ="EmptyCell977"
                                    GroupTable =32
                                    LayoutCachedLeft =6144
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =6468
                                    Width =2292
                                    Height =204
                                    Name ="EmptyCell1300"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =6468
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =6672
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =576
                                    Top =6732
                                    Width =1536
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label656"
                                    Caption ="LPAF Date"
                                    GroupTable =32
                                    LayoutCachedLeft =576
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2172
                                    Top =6732
                                    Width =1452
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label657"
                                    Caption ="BDEX"
                                    GroupTable =32
                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =3684
                                    Top =6732
                                    Width =912
                                    Name ="EmptyCell658"
                                    GroupTable =32
                                    LayoutCachedLeft =3684
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =4596
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4656
                                    Top =6732
                                    Width =1428
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label527"
                                    Caption ="ADMIN"
                                    GroupTable =32
                                    LayoutCachedLeft =4656
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =6084
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =6144
                                    Top =6732
                                    Width =1368
                                    Height =240
                                    FontSize =10
                                    ForeColor =9211020
                                    Name ="Label523"
                                    Caption ="ATF Date"
                                    GroupTable =32
                                    LayoutCachedLeft =6144
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    BorderShade =85.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =6732
                                    Width =2292
                                    Name ="EmptyCell1301"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =6732
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =6972
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =576
                                    Top =7032
                                    Width =1536
                                    Height =336
                                    FontWeight =700
                                    TabIndex =9
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="LPAFcomp3"
                                    ControlSource ="=CDate([OutcomeDate])"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =576
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =2112
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2172
                                    Top =7032
                                    Width =1452
                                    Height =336
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =10
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="BDEX3"
                                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                                    GroupTable =32

                                    LayoutCachedLeft =2172
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =3624
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =3684
                                    Top =7032
                                    Width =912
                                    Height =336
                                    Name ="EmptyCell773"
                                    GroupTable =32
                                    LayoutCachedLeft =3684
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =4596
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4656
                                    Top =7032
                                    Width =1428
                                    Height =336
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =11
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="ATFAdmin3"
                                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="Admin is Required"
                                    DefaultValue ="=[FAUserInit]"
                                    FontName ="Segoe UI"
                                    Tag ="Required"
                                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                                    GroupTable =32

                                    LayoutCachedLeft =4656
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =6084
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6144
                                    Top =7032
                                    Width =1368
                                    Height =336
                                    FontWeight =700
                                    TabIndex =12
                                    BorderColor =14277081
                                    ForeColor =4144959
                                    Name ="ATFDate3"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    GroupTable =32

                                    LayoutCachedLeft =6144
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =7512
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =3
                                    BorderShade =85.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GroupTable =32
                                End
                                Begin EmptyCell
                                    Left =7572
                                    Top =7032
                                    Width =2292
                                    Height =336
                                    Name ="EmptyCell1302"
                                    GroupTable =32
                                    LayoutCachedLeft =7572
                                    LayoutCachedTop =7032
                                    LayoutCachedWidth =9864
                                    LayoutCachedHeight =7368
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =3
                                    GroupTable =32
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =75
                    Width =720
                    Height =345
                    FontSize =12
                    TabIndex =4
                    ForeColor =8210719
                    Name ="StepLabel4"
                    ControlSource ="=\"Step 4\""

                    LayoutCachedLeft =2220
                    LayoutCachedTop =75
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =75
                    Width =720
                    Height =345
                    FontSize =12
                    TabIndex =3
                    ForeColor =8210719
                    Name ="StepLabel3"
                    ControlSource ="=\"Step 3\""

                    LayoutCachedLeft =1500
                    LayoutCachedTop =75
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =75
                    Width =720
                    Height =345
                    FontSize =12
                    TabIndex =2
                    ForeColor =8210719
                    Name ="StepLabel2"
                    ControlSource ="=\"Step 2\""

                    LayoutCachedLeft =780
                    LayoutCachedTop =75
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =75
                    Width =720
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8210719
                    Name ="StepLabel1"
                    ControlSource ="=\"Step 1\""

                    LayoutCachedLeft =60
                    LayoutCachedTop =75
                    LayoutCachedWidth =780
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9405
                    Width =840
                    Height =252
                    FontSize =9
                    TabIndex =9
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =9405
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6885
                    Width =840
                    Height =252
                    FontSize =9
                    TabIndex =8
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="OpportunityId"
                    ControlSource ="OpportunityId"

                    LayoutCachedLeft =6885
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6045
                    Width =840
                    Height =252
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =7
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="AccountID"
                    ControlSource ="AccountId"

                    LayoutCachedLeft =6045
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8565
                    Width =840
                    Height =252
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =6
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupID"

                    LayoutCachedLeft =8565
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7725
                    Width =840
                    Height =252
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =5
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="SpecId"
                    ControlSource ="SpecId"

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =8565
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5160
                    Width =840
                    Height =252
                    FontSize =9
                    TabIndex =10
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="SFAccountId"

                    LayoutCachedLeft =5160
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
            End
        End
        Begin FormFooter
            Height =840
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =5955
                    Top =120
                    Width =1155
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =6108695
                    Name ="GoNext"
                    Caption ="Next"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Proceed to Next Step"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5955
                    LayoutCachedTop =120
                    LayoutCachedWidth =7110
                    LayoutCachedHeight =570
                    PictureCaptionArrangement =4
                    Alignment =3
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    UseTheme =255
                    Shape =0
                    Gradient =0
                    BackColor =14270637
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =11525325
                    PressedColor =14211288
                    HoverForeColor =6108695
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    PictureType =1
                    Left =7290
                    Top =120
                    Width =2340
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =6108695
                    Name ="ImportToAtlaas"
                    Caption =" Import To ATLAAS"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    Tag ="Buttons"
                    ControlTipText ="Create the Schedule / Groups / Units into ATLAAS Core"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7290
                    LayoutCachedTop =120
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =570
                    PictureCaptionArrangement =5
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    UseTheme =255
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =11525325
                    PressedColor =14211288
                    HoverForeColor =6108695
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =4620
                    Top =120
                    Width =1155
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =6108695
                    Name ="GoBack"
                    Caption ="Back "
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Go Back to the Previous Step"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4620
                    LayoutCachedTop =120
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =570
                    PictureCaptionArrangement =5
                    Alignment =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    UseTheme =255
                    Shape =0
                    Gradient =0
                    BackColor =14270637
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =11525325
                    PressedColor =14211288
                    HoverForeColor =6108695
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =540
                    Top =90
                    Width =885
                    Height =240
                    FontSize =8
                    BackColor =10855845
                    BorderColor =0
                    Name ="SchedulesLabel"
                    Caption ="Schedules"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =540
                    LayoutCachedTop =90
                    LayoutCachedWidth =1425
                    LayoutCachedHeight =330
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =540
                    Top =285
                    Width =885
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9211020
                    ForeColor =7500402
                    Name ="NumOfSchedules"
                    ControlTipText ="Number on Schedules that will be inserted into ATLAAS Core"
                    GridlineColor =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =285
                    LayoutCachedWidth =1425
                    LayoutCachedHeight =645
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1545
                    Top =90
                    Width =885
                    Height =240
                    FontSize =8
                    BackColor =10855845
                    BorderColor =0
                    Name ="GroupsLabel"
                    Caption ="Groups"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1545
                    LayoutCachedTop =90
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =330
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1545
                    Top =285
                    Width =885
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9211020
                    ForeColor =7500402
                    Name ="NumOfGroups"
                    ControlTipText ="Number on Groups that will be inserted into ATLAAS Core"
                    GridlineColor =0

                    LayoutCachedLeft =1545
                    LayoutCachedTop =285
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =645
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =2550
                    Top =90
                    Width =885
                    Height =240
                    FontSize =8
                    BackColor =10855845
                    BorderColor =0
                    Name ="Label167"
                    Caption ="Units"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2550
                    LayoutCachedTop =90
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =330
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2550
                    Top =285
                    Width =885
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9211020
                    ForeColor =7500402
                    Name ="NumOfUnits"
                    ControlTipText ="Number on Units that will be inserted into ATLAAS Core"
                    GridlineColor =0

                    LayoutCachedLeft =2550
                    LayoutCachedTop =285
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =645
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3585
                    Top =345
                    Width =840
                    ColumnOrder =0
                    FontSize =9
                    BackColor =15921906
                    BorderColor =14277081
                    Name ="LesseeID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =3585
                    LayoutCachedTop =345
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =585
                    ColumnStart =1
                    ColumnEnd =1
                    BackShade =95.0
                    BorderShade =85.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
            End
        End
    End
End
CodeBehindForm
' See "ATF_SFImporter.cls"
