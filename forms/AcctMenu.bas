Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17583
    DatasheetFontHeight =11
    ItemSuffix =202
    Left =4890
    Top =7395
    Right =22470
    Bottom =15975
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x653373b939cae340
    End
    Caption ="Accounting Menu"
    OnClose ="[Event Procedure]"
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
        Begin OptionButton
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
            BorderColor =12632256
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =840
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =660
                    Width =17568
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =660
                    LayoutCachedWidth =17568
                    LayoutCachedHeight =840
                    TabIndex =4
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =15810
                    Top =165
                    Width =750
                    Height =228
                    FontSize =8
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15810
                    LayoutCachedTop =165
                    LayoutCachedWidth =16560
                    LayoutCachedHeight =393
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
                    Left =15255
                    Top =165
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15255
                    LayoutCachedTop =165
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =393
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
                Begin Label
                    OverlapFlags =255
                    TextAlign =3
                    Width =2040
                    Height =690
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Accounting"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =690
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Top =375
                    Width =1815
                    Height =240
                    FontSize =8
                    Name ="Label227"
                    Caption =" MENU"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =375
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6375
                    Top =120
                    Width =1080
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="EndDate"
                    Format ="Short Date"

                    LayoutCachedLeft =6375
                    LayoutCachedTop =120
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =420
                End
                Begin Label
                    OverlapFlags =247
                    Left =5340
                    Top =480
                    Width =1860
                    Height =225
                    FontSize =10
                    Name ="Label184"
                    Caption ="▲Start - <--> - End▲"
                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =705
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5175
                    Top =120
                    Width =1080
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="StartDate"
                    Format ="Short Date"

                    LayoutCachedLeft =5175
                    LayoutCachedTop =120
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =420
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7755
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =690
                    Top =2400
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="InvoiceClients"
                    Caption ="Invoice ('Bill') Select"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =690
                    LayoutCachedTop =2400
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2730
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =93
                    Left =210
                    Top =3570
                    Width =7455
                    Height =15
                    Name ="Line13"
                    LayoutCachedLeft =210
                    LayoutCachedTop =3570
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =3585
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =1
                    Left =8370
                    Top =135
                    Width =4275
                    Height =405
                    FontSize =14
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label14"
                    Caption ="Transaction Economic Summary"
                    LayoutCachedLeft =8370
                    LayoutCachedTop =135
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =10212
                    Top =492
                    Width =2616
                    Height =312
                    FontSize =10
                    TabIndex =2
                    BorderColor =12566463
                    Name ="SelClient"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName, ClientGroups.ClientGroup"
                        "Name FROM ClientGroups INNER JOIN vw_SixKeys ON ClientGroups.ClientGroupId = vw_"
                        "SixKeys.ClientGroupId WHERE (((vw_SixKeys.MLOrig)=\"FA\")) GROUP BY vw_SixKeys.c"
                        "lientID, vw_SixKeys.ClientShortName, ClientGroups.ClientGroupName ORDER BY vw_Si"
                        "xKeys.ClientShortName;"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =10212
                    LayoutCachedTop =492
                    LayoutCachedWidth =12828
                    LayoutCachedHeight =804
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =8688
                            Top =492
                            Width =1452
                            Height =312
                            FontSize =10
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="ClientShNm_Label"
                            Caption ="1. Customer"
                            GroupTable =1
                            LayoutCachedLeft =8688
                            LayoutCachedTop =492
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =804
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =10212
                    Top =876
                    Width =2616
                    Height =312
                    FontSize =10
                    TabIndex =3
                    BorderColor =12566463
                    Name ="SelLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLNo FROM MstrLease WHERE (((MstrLease.MLOrig)="
                        "\"FA\") AND ((MstrLease.Client)=[forms]![AcctMenu]![selclient])) ORDER BY MstrLe"
                        "ase.MLNo;"
                    ColumnWidths ="0;1440"
                    StatusBarText ="Only Master Leases with Orig FA appear in this list."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Only Master Leases with Orig FA appear in this list."
                    GroupTable =1

                    LayoutCachedLeft =10212
                    LayoutCachedTop =876
                    LayoutCachedWidth =12828
                    LayoutCachedHeight =1188
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8688
                            Top =876
                            Width =1452
                            Height =312
                            FontSize =10
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label4"
                            Caption ="2. Lease"
                            GroupTable =1
                            LayoutCachedLeft =8688
                            LayoutCachedTop =876
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =1188
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =10212
                    Top =1644
                    Width =2616
                    Height =312
                    FontSize =10
                    TabIndex =5
                    BorderColor =12566463
                    ForeColor =-2147483617
                    Name ="SelGrp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchGrp.SGrpID, SchGrp.SGrpNO FROM (MstrLease INNER JOIN Schedule ON MstrL"
                        "ease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrp"
                        "SchID GROUP BY SchGrp.SGrpID, SchGrp.SGrpNO, Schedule.SchNo, IIf([schfsl]=True,\""
                        "yes\",Null), MstrLease.MLNo, SchGrp.SGrpSchID HAVING (((SchGrp.SGrpSchID)=[forms"
                        "]![AcctMenu]![selsch])) ORDER BY Schedule.SchNo DESC;"
                    ColumnWidths ="0;432"
                    GroupTable =1

                    LayoutCachedLeft =10212
                    LayoutCachedTop =1644
                    LayoutCachedWidth =12828
                    LayoutCachedHeight =1956
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8688
                            Top =1644
                            Width =1452
                            Height =312
                            FontSize =10
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label16"
                            Caption ="4. Group"
                            GroupTable =1
                            LayoutCachedLeft =8688
                            LayoutCachedTop =1644
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =1956
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8688
                    Top =2028
                    Width =4140
                    Height =312
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="AddExpBtn"
                    Caption ="Group Expenses"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =8688
                    LayoutCachedTop =2028
                    LayoutCachedWidth =12828
                    LayoutCachedHeight =2340
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10200
                    Top =7140
                    Width =3555
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =31
                    Name ="TESRptBtn"
                    Caption ="Transaction Economic Summary"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10200
                    LayoutCachedTop =7140
                    LayoutCachedWidth =13755
                    LayoutCachedHeight =7455
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8685
                    Top =7125
                    Width =1230
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label24"
                    Caption ="9. Run Report"
                    LayoutCachedLeft =8685
                    LayoutCachedTop =7125
                    LayoutCachedWidth =9915
                    LayoutCachedHeight =7410
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =8655
                    Top =2460
                    Width =4125
                    Height =1215
                    TabIndex =15
                    Name ="CommSubForm"
                    SourceObject ="Form.CommSubForm"
                    LinkChildFields ="SchID"
                    LinkMasterFields ="SelSch"

                    LayoutCachedLeft =8655
                    LayoutCachedTop =2460
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =3675
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8685
                    Top =3690
                    Width =1305
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label33"
                    Caption ="7. Prepayment"
                    LayoutCachedLeft =8685
                    LayoutCachedTop =3690
                    LayoutCachedWidth =9990
                    LayoutCachedHeight =3975
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =8865
                    Top =4005
                    Width =4110
                    Height =2085
                    TabIndex =22
                    Name ="CommSubF"
                    SourceObject ="Form.CommSubF"

                    LayoutCachedLeft =8865
                    LayoutCachedTop =4005
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =6090
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =705
                    Top =600
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    Name ="SchTrk"
                    Caption ="Syndication Readiness  (PDF) Report"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =705
                    LayoutCachedTop =600
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =930
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =93
                    Left =210
                    Top =375
                    Width =7455
                    Height =15
                    Name ="Line40"
                    LayoutCachedLeft =210
                    LayoutCachedTop =375
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =390
                    BorderThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4050
                    Top =600
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Command42"
                    Caption ="Client Review Report"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="AcctClientReview"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command42\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argume"
                                "nt Name=\"ReportName\">AcctClient"
                        End
                        Begin
                            Comment ="_AXL:Review</Argument><Argument Name=\"View\">Print Preview</Argument></Action><"
                                "/Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =600
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =930
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =705
                    Top =1065
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="LoanReport"
                    Caption ="Loan Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =705
                    LayoutCachedTop =1065
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    OverlapFlags =85
                    Left =7935
                    Width =15
                    Height =7755
                    BorderColor =8872517
                    Name ="Line47"
                    LayoutCachedLeft =7935
                    LayoutCachedWidth =7950
                    LayoutCachedHeight =7755
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4050
                    Top =2400
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    Name ="PotentialBillingExcel"
                    Caption ="  Potential Billing"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =2400
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =2730
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4050
                    Top =7200
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    Name ="Command54"
                    Caption ="Employee  Maintenance "
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Employee Dropdown Maint"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command54\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">Employee Dropd"
                        End
                        Begin
                            Comment ="_AXL:own Maint</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =7200
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =7530
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15516
                    Top =492
                    Width =1080
                    Height =312
                    FontSize =10
                    TabIndex =7
                    BorderColor =12566463
                    Name ="currrt"
                    Format ="Percent"
                    DefaultValue ="0"
                    GroupTable =2

                    LayoutCachedLeft =15516
                    LayoutCachedTop =492
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =804
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13248
                            Top =492
                            Width =2196
                            Height =312
                            FontSize =10
                            FontWeight =700
                            Name ="Label65"
                            Caption ="Current 1mth Libor Rate:"
                            GroupTable =2
                            LayoutCachedLeft =13248
                            LayoutCachedTop =492
                            LayoutCachedWidth =15444
                            LayoutCachedHeight =804
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15516
                    Top =876
                    Width =1080
                    Height =312
                    FontSize =10
                    TabIndex =8
                    BorderColor =12566463
                    Name ="spread"
                    Format ="Percent"
                    DefaultValue ="0"
                    GroupTable =2

                    LayoutCachedLeft =15516
                    LayoutCachedTop =876
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =1188
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13248
                            Top =876
                            Width =2196
                            Height =312
                            FontSize =10
                            FontWeight =700
                            Name ="Label67"
                            Caption ="Spread:"
                            GroupTable =2
                            LayoutCachedLeft =13248
                            LayoutCachedTop =876
                            LayoutCachedWidth =15444
                            LayoutCachedHeight =1188
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =13065
                    Top =465
                    Width =3975
                    Height =1170
                    Name ="Box68"
                    LayoutCachedLeft =13065
                    LayoutCachedTop =465
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =1635
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15516
                    Top =1260
                    Width =1080
                    Height =312
                    FontSize =10
                    TabIndex =9
                    BorderColor =12566463
                    Name ="total"
                    ControlSource ="=[currrt]+[spread]"
                    Format ="Percent"
                    GroupTable =2

                    LayoutCachedLeft =15516
                    LayoutCachedTop =1260
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =1572
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =13065
                    Top =135
                    Width =3795
                    Height =405
                    FontSize =14
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label74"
                    Caption ="Customer Interest Charge Info"
                    LayoutCachedLeft =13065
                    LayoutCachedTop =135
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15516
                    Top =2172
                    Width =1080
                    Height =312
                    TabIndex =10
                    BorderColor =12566463
                    Name ="LineRt"
                    Format ="Percent"
                    DefaultValue ="0"
                    GroupTable =2

                    LayoutCachedLeft =15516
                    LayoutCachedTop =2172
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =2484
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13248
                            Top =2172
                            Width =2196
                            Height =312
                            Name ="Label76"
                            Caption =" Libor Rate:"
                            GroupTable =2
                            LayoutCachedLeft =13248
                            LayoutCachedTop =2172
                            LayoutCachedWidth =15444
                            LayoutCachedHeight =2484
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =13248
                    Top =1764
                    Width =3348
                    Height =336
                    FontSize =14
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label77"
                    Caption ="Interest Expense"
                    GroupTable =2
                    LayoutCachedLeft =13248
                    LayoutCachedTop =1764
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =2100
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4050
                    Top =6705
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =27
                    Name ="OpenClientsForm"
                    Caption ="All Client/Target Lookup"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4050
                    LayoutCachedTop =6705
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =7035
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =690
                    Top =7200
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =29
                    Name ="UlookBtn"
                    Caption ="Unit Lookup"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =690
                    LayoutCachedTop =7200
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =7530
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =720
                    Top =4260
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    Name ="ManageInvoice"
                    Caption ="Create / Pay Invoice"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =4260
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =4590
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =720
                    Top =5370
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =23
                    Name ="exchbtn"
                    Caption ="EXchangeIT Monitor"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="ExChRpt"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"exchbtn\" xmlns=\"http://schemas.microsoft.com/office/access"
                                "services/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument N"
                                "ame=\"FormName\">ExChRpt</Argumen"
                        End
                        Begin
                            Comment ="_AXL:t></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =5370
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =5700
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4080
                    Top =4920
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =24
                    Name ="CasualtyReport"
                    Caption ="Casualty Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =2

                    LayoutCachedLeft =4080
                    LayoutCachedTop =4920
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =5250
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2730
                    Top =3780
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    Name ="Command12"
                    Caption ="Remarketing Menu"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OpenForm"
                            Argument ="Rmkt Menu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command12\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">Rmkt Menu</Arg"
                        End
                        Begin
                            Comment ="_AXL:ument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2730
                    LayoutCachedTop =3780
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =4110
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8685
                    Top =6405
                    Width =1560
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label94"
                    Caption ="8. Override BD EX"
                    LayoutCachedLeft =8685
                    LayoutCachedTop =6405
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =6690
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10200
                    Top =6360
                    Width =2610
                    Height =315
                    TabIndex =28
                    BorderColor =12566463
                    Name ="LOExOV"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Salesperson.SalespersonName FROM Salesperson WHERE (((Salesperson.Type)=\""
                        "1\") AND ((Salesperson.SalespersonInactive)=No)) ORDER BY Salesperson.SPInitials"
                        ";"
                    ColumnWidths ="2160"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =6360
                    LayoutCachedWidth =12810
                    LayoutCachedHeight =6675
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =690
                    Top =6690
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =26
                    Name ="Command48"
                    Caption ="Client/Target Search"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Search_clients"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command48\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">Search_clients"
                        End
                        Begin
                            Comment ="_AXL:</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =690
                    LayoutCachedTop =6690
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =7020
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =720
                    Top =4920
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    Name ="RemarkingBuyOut"
                    Caption =" Purchase / Unit Buy Out"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =4920
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =5250
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4080
                    Top =4260
                    Width =2010
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    Name ="RemarketingInvReport"
                    Caption ="Inventory"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4080
                    LayoutCachedTop =4260
                    LayoutCachedWidth =6090
                    LayoutCachedHeight =4590
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Top =240
                    Width =1545
                    Height =300
                    FontWeight =700
                    Name ="Label110"
                    Caption ="Lease Reports"
                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =540
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =93
                    Left =210
                    Top =1725
                    Width =7455
                    Height =15
                    Name ="Line111"
                    LayoutCachedLeft =210
                    LayoutCachedTop =1725
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =1740
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Top =1545
                    Width =1545
                    Height =300
                    FontWeight =700
                    Name ="Label112"
                    Caption ="Lease Invoicing"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1545
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =1845
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Top =3405
                    Width =1545
                    Height =300
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Remarketing"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3405
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =3705
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =6090
                    Top =4260
                    Width =1155
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    Name ="Command116"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6090
                    LayoutCachedTop =4260
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =4590
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6135
                    Top =4320
                    Width =1005
                    Height =210
                    FontSize =9
                    TabIndex =21
                    Name ="AsOfDate"
                    Format ="Short Date"

                    LayoutCachedLeft =6135
                    LayoutCachedTop =4320
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =4530
                End
                Begin Label
                    OverlapFlags =215
                    Left =6030
                    Top =4620
                    Width =1185
                    Height =225
                    FontSize =9
                    Name ="Label117"
                    Caption ="'As Of'▲Date"
                    LayoutCachedLeft =6030
                    LayoutCachedTop =4620
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =4845
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =93
                    Left =210
                    Top =6375
                    Width =7455
                    Height =15
                    Name ="Line118"
                    LayoutCachedLeft =210
                    LayoutCachedTop =6375
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =6390
                    BorderThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Top =6210
                    Width =1545
                    Height =300
                    FontWeight =700
                    Name ="Label119"
                    Caption ="Look Ups"
                    LayoutCachedLeft =60
                    LayoutCachedTop =6210
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =6510
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =0
                End
                Begin Label
                    Vertical = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    Left =8355
                    Top =525
                    Width =345
                    Height =6105
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label120"
                    Caption ="└────────────────────── Select ─────────────────────┘"
                    LayoutCachedLeft =8355
                    LayoutCachedTop =525
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =6630
                End
                Begin EmptyCell
                    Left =13248
                    Top =1260
                    Width =2196
                    Height =312
                    Name ="EmptyCell173"
                    GroupTable =2
                    LayoutCachedLeft =13248
                    LayoutCachedTop =1260
                    LayoutCachedWidth =15444
                    LayoutCachedHeight =1572
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =13248
                    Top =1632
                    Width =3348
                    Height =72
                    Name ="EmptyCell178"
                    GroupTable =2
                    LayoutCachedLeft =13248
                    LayoutCachedTop =1632
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =1704
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =13065
                    Top =2085
                    Width =3975
                    Height =600
                    Name ="Box180"
                    LayoutCachedLeft =13065
                    LayoutCachedTop =2085
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =2685
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =4080
                    Top =5370
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =25
                    Name ="ReserveReport"
                    Caption ="Reserve"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4080
                    LayoutCachedTop =5370
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =5700
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedThemeColorIndex =7
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4050
                    Top =1065
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="Assignee2Client"
                    Caption ="Assignee to Client Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =1065
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =2880
                    Left =10212
                    Top =1260
                    Width =2616
                    Height =312
                    FontSize =10
                    TabIndex =4
                    BorderColor =12566463
                    Name ="SelSch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.SchID, vw_SixKeys.Schedule, Max(vw_SixKeys.BLCD) AS BLCD, IIf("
                        "[schfsl]=True,\"yes\",Null) AS Expr1\015\012FROM Schedule INNER JOIN vw_SixKeys "
                        "ON Schedule.SchID = vw_SixKeys.SchID\015\012WHERE (((vw_SixKeys.MLID)=[forms]![A"
                        "cctMenu]![selLease]))\015\012GROUP BY vw_SixKeys.SchID, vw_SixKeys.Schedule, IIf"
                        "([schfsl]=True,\"yes\",Null)\015\012ORDER BY vw_SixKeys.SchID DESC;"
                    ColumnWidths ="0;1440;1008;432"
                    StatusBarText ="Only Master Leases with Orig FA appear in this list."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Only Master Leases with Orig FA appear in this list."
                    GroupTable =1

                    LayoutCachedLeft =10212
                    LayoutCachedTop =1260
                    LayoutCachedWidth =12828
                    LayoutCachedHeight =1572
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =8688
                            Top =1260
                            Width =1452
                            Height =312
                            FontSize =10
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label194"
                            Caption ="3. Schedule"
                            GroupTable =1
                            LayoutCachedLeft =8688
                            LayoutCachedTop =1260
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =1572
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =4050
                    Top =1950
                    Width =3150
                    Height =330
                    FontSize =9
                    TabIndex =32
                    Name ="SurrenderPDReport"
                    Caption =" Surrender / HoldOver PerDiem"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =1950
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =2280
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =7
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedThemeColorIndex =7
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =33
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =4050
                    Top =2850
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =33
                    Name =" InvoicedReport"
                    Caption ="  Invoiced Report"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Ctl_InvoicedReport"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =4050
                    LayoutCachedTop =2850
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =3180
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =720
                    Top =2850
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =34
                    Name ="PotentialBillingPDExcel"
                    Caption ="  Potential Billing PERDIEM"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =2850
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =3180
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =690
                    Top =1950
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =35
                    Name ="BatchBillingForm"
                    Caption ="Batch Billing Manager"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =690
                    LayoutCachedTop =1950
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2280
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =5676533
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =5676533
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4080
                    Top =5850
                    Width =3150
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =36
                    Name ="InterimStatusReport"
                    Caption ="  Interim Status Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =4080
                    LayoutCachedTop =5850
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =6180
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
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
' See "AcctMenu.cls"
