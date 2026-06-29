Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =19257
    ItemSuffix =611
    Left =8550
    Top =8400
    Right =27855
    Bottom =16410
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x3572a1eb77c8e440
    End
    RecordSource ="SELECT MstrLease.*, Clients.clientTaxID FROM MstrLease LEFT JOIN Clients ON Mstr"
        "Lease.MLGuarantor = Clients.clientID ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Moveable =0
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
            Height =630
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1335
                    Width =1065
                    Height =0
                    ColumnOrder =0
                    Name ="clientID"
                    ControlSource ="client"

                    LayoutCachedLeft =1335
                    LayoutCachedWidth =2400
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Width =19257
                    Height =480
                    FontSize =18
                    ForeColor =10040879
                    Name ="Label35"
                    Caption ="MASTER LEASES"
                    FontName ="Segoe UI"
                    HorizontalAnchor =2
                    LayoutCachedWidth =19257
                    LayoutCachedHeight =480
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =15375
                    Top =105
                    Width =1815
                    Height =315
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="ADDML"
                    Caption ="Add a Master Lease"
                    OnClick ="[Event Procedure]"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =120

                    LayoutCachedLeft =15375
                    LayoutCachedTop =105
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =16777215
                    HoverThemeColorIndex =1
                    HoverShade =50.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureAlignment =0
                    PictureType =2
                    Top =450
                    Width =19257
                    Height =180
                    Name ="Image4124"
                    Picture ="LtGrey2_Shadow"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =450
                    LayoutCachedWidth =19257
                    LayoutCachedHeight =630
                    TabIndex =2
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7260
            Name ="Detail"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =2145
                    Top =165
                    Width =7830
                    Height =1200
                    Name ="Box476"
                    LayoutCachedLeft =2145
                    LayoutCachedTop =165
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =1365
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =2145
                    Top =1605
                    Width =7830
                    Height =1530
                    Name ="Box475"
                    LayoutCachedLeft =2145
                    LayoutCachedTop =1605
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =3135
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =2145
                    Top =3420
                    Width =14775
                    Height =900
                    Name ="Box415"
                    LayoutCachedLeft =2145
                    LayoutCachedTop =3420
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =4320
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3408
                    Top =1020
                    Width =1200
                    Height =252
                    FontSize =8
                    TabIndex =5
                    Name ="LeaseDate"
                    ControlSource ="MLDT"
                    FontName ="Segoe UI"
                    GroupTable =12
                    BottomPadding =150

                    LayoutCachedLeft =3408
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =2460
                            Top =1020
                            Width =870
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label253"
                            Caption ="Lease Date"
                            FontName ="Segoe UI"
                            GroupTable =12
                            BottomPadding =150
                            LayoutCachedLeft =2460
                            LayoutCachedTop =1020
                            LayoutCachedWidth =3330
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6576
                    Top =2124
                    Width =2928
                    Height =252
                    FontSize =8
                    TabIndex =20
                    Name ="LeaseMLASigTitle"
                    ControlSource ="MLSignertitle"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6576
                    LayoutCachedTop =2124
                    LayoutCachedWidth =9504
                    LayoutCachedHeight =2376
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12048
                    Top =1224
                    Width =4812
                    Height =252
                    FontSize =8
                    TabIndex =11
                    Name ="LeaseGtySigTitle"
                    ControlSource ="MLGuarSignTitle"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =1224
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =1476
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =1224
                            Width =1608
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label262"
                            Caption ="Signer's Title"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =1224
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =1476
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12048
                    Top =1548
                    Width =1932
                    Height =252
                    FontSize =8
                    TabIndex =12
                    Name ="LeaseGtyTaxID"
                    ControlSource ="clienttaxid"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =1548
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =1548
                            Width =1608
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label263"
                            Caption ="Guarantor Tax ID"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =1548
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =1800
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6576
                    Top =2448
                    Width =2928
                    Height =252
                    FontSize =8
                    TabIndex =23
                    Name ="MLSignerTitle2"
                    ControlSource ="MLSignerTitle2"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6576
                    LayoutCachedTop =2448
                    LayoutCachedWidth =9504
                    LayoutCachedHeight =2700
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6576
                    Top =2772
                    Width =2928
                    Height =252
                    FontSize =8
                    TabIndex =26
                    Name ="Text9"
                    ControlSource ="MLSignerTitle3"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6576
                    LayoutCachedTop =2772
                    LayoutCachedWidth =9504
                    LayoutCachedHeight =3024
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5508
                    Top =1020
                    Width =1200
                    Height =252
                    FontSize =8
                    TabIndex =6
                    Name ="MLSent"
                    ControlSource ="MLSent"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =12
                    BottomPadding =150

                    LayoutCachedLeft =5508
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6708
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4668
                            Top =1020
                            Width =768
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label11"
                            Caption ="Sent"
                            FontName ="Segoe UI"
                            GroupTable =12
                            BottomPadding =150
                            LayoutCachedLeft =4668
                            LayoutCachedTop =1020
                            LayoutCachedWidth =5436
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7680
                    Top =1020
                    Width =1200
                    Height =252
                    FontSize =8
                    TabIndex =7
                    Name ="MLBack"
                    ControlSource ="MLBack"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =12
                    BottomPadding =150

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =6780
                            Top =1020
                            Width =828
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label12"
                            Caption ="Returned"
                            FontName ="Segoe UI"
                            GroupTable =12
                            BottomPadding =150
                            LayoutCachedLeft =6780
                            LayoutCachedTop =1020
                            LayoutCachedWidth =7608
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =2280
                    Top =3210
                    Width =990
                    Height =345
                    Name ="Label13"
                    Caption ="Insurance"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2280
                    LayoutCachedTop =3210
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =3555
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5988
                    Top =3624
                    Width =1560
                    FontSize =8
                    TabIndex =29
                    BorderColor =8872517
                    Name ="MLINReq"
                    ControlSource ="MLINReq"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Minimum Insurance Amount Required"
                    ValidationRule =">=0"
                    ValidationText ="Liability Amt Req must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =5988
                    LayoutCachedTop =3624
                    LayoutCachedWidth =7548
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =14256
                    Top =3624
                    Width =756
                    TabIndex =35
                    Name ="MLStipLoss"
                    ControlSource ="MLStipLoss"
                    StatusBarText ="is a Stip Loss table require in each lease"
                    GroupTable =29

                    LayoutCachedLeft =14256
                    LayoutCachedTop =3624
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Top =3624
                    Width =1560
                    FontSize =8
                    TabIndex =30
                    BorderColor =8872517
                    Name ="MLINRecA"
                    ControlSource ="MLINRecA"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="insurance amount received"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =7620
                    LayoutCachedTop =3624
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9252
                    Top =3624
                    Width =1200
                    FontSize =8
                    TabIndex =31
                    BorderColor =8872517
                    Name ="MLInsRec"
                    ControlSource ="MLInsRec"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =9252
                    LayoutCachedTop =3624
                    LayoutCachedWidth =10452
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =13176
                    Top =3624
                    Width =552
                    TabIndex =34
                    Name ="MLINSBLANKET"
                    ControlSource ="MLINSBLANKET"
                    StatusBarText ="IF YES AT LEASE LEVEL IS NO AT SCHEDULE LEVEL"
                    GroupTable =29

                    LayoutCachedLeft =13176
                    LayoutCachedTop =3624
                    LayoutCachedWidth =13728
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4380
                    Top =3624
                    Width =1536
                    FontSize =8
                    TabIndex =28
                    BorderColor =8872517
                    Name ="MLPhysIN"
                    ControlSource ="MLLiabIns"
                    RowSourceType ="Value List"
                    RowSource ="S;self-insured;R;Required"
                    ColumnWidths ="0;1440"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3624
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =3432
                            Top =3624
                            Width =876
                            Height =240
                            FontSize =8
                            Name ="Label24"
                            Caption ="Liability"
                            FontName ="Segoe UI"
                            GroupTable =29
                            LayoutCachedLeft =3432
                            LayoutCachedTop =3624
                            LayoutCachedWidth =4308
                            LayoutCachedHeight =3864
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
                            BackThemeColorIndex =8
                            BackTint =40.0
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =29
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =12048
                    Top =600
                    Width =4812
                    Height =225
                    ColumnWidth =3300
                    FontSize =8
                    TabIndex =8
                    Name ="MLGuarantor"
                    ControlSource ="MLGuarantor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.clientID, Clients.clientCompanyName FROM Clients WHERE (((Clients"
                        ".clientType)=\"client\" Or (Clients.clientType)=\"guarantor\")) ORDER BY Clients"
                        ".clientCompanyName, Clients.clientType DESC;"
                    ColumnWidths ="0;2880"
                    StatusBarText ="ref to client table (guarantors only)"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =18
                    AllowValueListEdits =0

                    LayoutCachedLeft =12048
                    LayoutCachedTop =600
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =825
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =600
                            Width =1608
                            Height =225
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label25"
                            Caption ="Guarantor"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =600
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =825
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14052
                    Top =1872
                    Width =1080
                    Height =252
                    TabIndex =14
                    Name ="MLGuarUnlim"
                    ControlSource ="MLGuarUnlim"
                    StatusBarText ="Unlimited Guaranty?"
                    GroupTable =18

                    LayoutCachedLeft =14052
                    LayoutCachedTop =1872
                    LayoutCachedWidth =15132
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =13188
                            Top =1872
                            Width =792
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label26"
                            Caption ="Unlimited"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =13188
                            LayoutCachedTop =1872
                            LayoutCachedWidth =13980
                            LayoutCachedHeight =2124
                            RowStart =4
                            RowEnd =4
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12048
                    Top =1872
                    Width =1080
                    Height =252
                    FontSize =8
                    TabIndex =13
                    Name ="MLGuarLimit"
                    ControlSource ="MLGuarLimit"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText =" Limit must by >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =1872
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =1872
                            Width =1608
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label27"
                            Caption ="Guaranty Limit"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =1872
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =2124
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12048
                    Top =2520
                    Width =1080
                    Height =288
                    FontSize =8
                    TabIndex =17
                    Name ="MLGuarExp"
                    ControlSource ="MLGuarExp"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =2808
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =2520
                            Width =1608
                            Height =288
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label28"
                            Caption ="Guaranty Expiration"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =2520
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =2808
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12048
                    Top =2196
                    Width =1080
                    Height =252
                    FontSize =8
                    TabIndex =15
                    Name ="MLGuarSent"
                    ControlSource ="MLGuarSent"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =2196
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =2448
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10368
                            Top =2196
                            Width =1608
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label29"
                            Caption ="Guaranty Sent"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =2196
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =2448
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14052
                    Top =2196
                    Width =1080
                    Height =252
                    FontSize =8
                    TabIndex =16
                    Name ="MLGuarBack"
                    ControlSource ="MLGuarBack"
                    FontName ="Segoe UI"
                    GroupTable =18

                    LayoutCachedLeft =14052
                    LayoutCachedTop =2196
                    LayoutCachedWidth =15132
                    LayoutCachedHeight =2448
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13188
                            Top =2196
                            Width =792
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label30"
                            Caption ="Returned"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =13188
                            LayoutCachedTop =2196
                            LayoutCachedWidth =13980
                            LayoutCachedHeight =2448
                            RowStart =5
                            RowEnd =5
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6192
                    Top =1800
                    Width =312
                    Height =255
                    TabIndex =18
                    Name ="MLAnyCoASign"
                    ControlSource ="MLAnyCoASign"
                    StatusBarText ="if No signer must be incumbent"
                    GroupTable =1

                    LayoutCachedLeft =6192
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =2055
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4452
                            Top =1800
                            Width =1680
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label31"
                            Caption ="Any Signer-CofA"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =4452
                            LayoutCachedTop =1800
                            LayoutCachedWidth =6132
                            LayoutCachedHeight =2055
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2280
                    Top =30
                    Width =1080
                    Height =345
                    Name ="Label258"
                    Caption ="Lease Info"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2280
                    LayoutCachedTop =30
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =375
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3408
                    Top =372
                    Width =2688
                    Height =252
                    FontSize =8
                    TabIndex =1
                    BackColor =11271167
                    BorderColor =8872517
                    Name ="Text251"
                    ControlSource ="MLno"
                    FontName ="Segoe UI"
                    GroupTable =12

                    LayoutCachedLeft =3408
                    LayoutCachedTop =372
                    LayoutCachedWidth =6096
                    LayoutCachedHeight =624
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =2460
                            Top =372
                            Width =870
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label252"
                            Caption ="MLA No."
                            FontName ="Segoe UI"
                            GroupTable =12
                            LayoutCachedLeft =2460
                            LayoutCachedTop =372
                            LayoutCachedWidth =3330
                            LayoutCachedHeight =624
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =4
                    ListWidth =4320
                    Left =6780
                    Top =372
                    Width =2100
                    Height =252
                    FontSize =8
                    TabIndex =2
                    BackColor =11271167
                    BorderColor =8872517
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    RowSourceType ="Value List"
                    RowSource ="FA;FA Lease;Other;Other Leasing Entity;Rmkt;Units Purchased for Remarketing;Port"
                        "folio;C/P owned Units"
                    ColumnWidths ="1080;3240"
                    DefaultValue ="\"FA\""
                    FontName ="Segoe UI"
                    GroupTable =12
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =6780
                    LayoutCachedTop =372
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =624
                    ColumnStart =5
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =6156
                            Top =372
                            Width =552
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label2"
                            Caption ="Orig:"
                            FontName ="Segoe UI"
                            GroupTable =12
                            LayoutCachedLeft =6156
                            LayoutCachedTop =372
                            LayoutCachedWidth =6708
                            LayoutCachedHeight =624
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =12204
                    Top =3624
                    Width =660
                    TabIndex =33
                    Name ="MLINSAccpt"
                    ControlSource ="MLINSAccpt"
                    GroupTable =29

                    LayoutCachedLeft =12204
                    LayoutCachedTop =3624
                    LayoutCachedWidth =12864
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10524
                    Top =3624
                    Width =1200
                    FontSize =8
                    TabIndex =32
                    BorderColor =8872517
                    Name ="Text41"
                    ControlSource ="MLInsExp"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =10524
                    LayoutCachedTop =3624
                    LayoutCachedWidth =11724
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Top =3936
                    Width =1560
                    FontSize =8
                    TabIndex =37
                    BorderColor =8872517
                    Name ="MLPhyRec"
                    ControlSource ="MLPhyRec"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =7620
                    LayoutCachedTop =3936
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3408
                    Top =696
                    Width =5472
                    Height =252
                    FontSize =8
                    TabIndex =3
                    Name ="MLLessor"
                    ControlSource ="MLLessor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm FROM Clients WHERE (((Clients.clientType)=\"bank\") AN"
                        "D ((Clients.ClientShNm)<>\"lessor1\" And (Clients.ClientShNm)<>\"lessor2\")) ORD"
                        "ER BY Clients.ClientShNm;"
                    StatusBarText ="If other name of lessor"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =12
                    AllowValueListEdits =0

                    LayoutCachedLeft =3408
                    LayoutCachedTop =696
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =948
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =2460
                            Top =696
                            Width =870
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label3"
                            Caption ="Lessor"
                            FontName ="Segoe UI"
                            GroupTable =12
                            LayoutCachedLeft =2460
                            LayoutCachedTop =696
                            LayoutCachedWidth =3330
                            LayoutCachedHeight =948
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =12
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4380
                    Top =3936
                    Width =1536
                    FontSize =8
                    TabIndex =36
                    BorderColor =8872517
                    Name ="Combo54"
                    ControlSource ="MLPhysIN"
                    RowSourceType ="Value List"
                    RowSource ="S;self-insured;R;Required"
                    ColumnWidths ="0;1440"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3936
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =3432
                            Top =3936
                            Width =876
                            Height =240
                            FontSize =8
                            Name ="Label55"
                            Caption ="Physical"
                            FontName ="Segoe UI"
                            GroupTable =29
                            LayoutCachedLeft =3432
                            LayoutCachedTop =3936
                            LayoutCachedWidth =4308
                            LayoutCachedHeight =4176
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            BackThemeColorIndex =9
                            BackTint =40.0
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =29
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =12204
                    Top =3936
                    Width =660
                    TabIndex =39
                    Name ="PhysIncAccpt"
                    ControlSource ="PhysIncAccpt"
                    GroupTable =29

                    LayoutCachedLeft =12204
                    LayoutCachedTop =3936
                    LayoutCachedWidth =12864
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4380
                    Top =4248
                    Width =10632
                    FontSize =8
                    TabIndex =40
                    BorderColor =10040879
                    Name ="INScomment"
                    ControlSource ="INScomment"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =4380
                    LayoutCachedTop =4248
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =4488
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =11
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =3432
                            Top =4248
                            Width =876
                            Height =240
                            FontSize =8
                            Name ="Label57"
                            Caption ="Comment"
                            FontName ="Segoe UI"
                            GroupTable =29
                            LayoutCachedLeft =3432
                            LayoutCachedTop =4248
                            LayoutCachedWidth =4308
                            LayoutCachedHeight =4488
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =4
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =29
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10524
                    Top =3936
                    Width =1200
                    FontSize =8
                    TabIndex =38
                    BorderColor =8872517
                    Name ="PhysInsExp"
                    ControlSource ="PhysInsExp"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =29

                    LayoutCachedLeft =10524
                    LayoutCachedTop =3936
                    LayoutCachedWidth =11724
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =2070
                    Top =4470
                    Width =15285
                    Height =2760
                    TabIndex =42
                    Name ="MLAmendf"
                    SourceObject ="Form.MLAmendf"
                    LinkChildFields ="MLRef"
                    LinkMasterFields ="MLID"

                    LayoutCachedLeft =2070
                    LayoutCachedTop =4470
                    LayoutCachedWidth =17355
                    LayoutCachedHeight =7230
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =6576
                    Top =1800
                    Width =2928
                    Height =255
                    FontSize =8
                    ForeColor =10040879
                    Name ="Label4"
                    Caption ="Title"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =6576
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9504
                    LayoutCachedHeight =2055
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =3420
                    Top =1800
                    Width =972
                    Height =255
                    FontSize =8
                    ForeColor =10040879
                    Name ="Label91"
                    Caption ="Full Name"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =3420
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4392
                    LayoutCachedHeight =2055
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2472
                    Top =1800
                    Width =876
                    Height =255
                    Name ="EmptyCell94"
                    GroupTable =1
                    LayoutCachedLeft =2472
                    LayoutCachedTop =1800
                    LayoutCachedWidth =3348
                    LayoutCachedHeight =2055
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =2280
                    Top =1410
                    Width =765
                    Height =345
                    Name ="Label209"
                    Caption ="Signers"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2280
                    LayoutCachedTop =1410
                    LayoutCachedWidth =3045
                    LayoutCachedHeight =1755
                    BackThemeColorIndex =1
                End
                Begin EmptyCell
                    Left =5988
                    Top =3936
                    Width =1560
                    Name ="EmptyCell331"
                    GroupTable =29
                    LayoutCachedLeft =5988
                    LayoutCachedTop =3936
                    LayoutCachedWidth =7548
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =9252
                    Top =3936
                    Width =1200
                    Name ="EmptyCell335"
                    GroupTable =29
                    LayoutCachedLeft =9252
                    LayoutCachedTop =3936
                    LayoutCachedWidth =10452
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =10524
                    Top =3276
                    Width =1200
                    Height =276
                    FontSize =8
                    Name ="Label356"
                    Caption ="Expires "
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =10524
                    LayoutCachedTop =3276
                    LayoutCachedWidth =11724
                    LayoutCachedHeight =3552
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =3432
                    Top =3276
                    Width =876
                    Height =276
                    Name ="EmptyCell359"
                    GroupTable =29
                    LayoutCachedLeft =3432
                    LayoutCachedTop =3276
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =3552
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =9252
                    Top =3276
                    Width =1200
                    Height =276
                    FontSize =8
                    Name ="Label21"
                    Caption ="Date Recv"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =9252
                    LayoutCachedTop =3276
                    LayoutCachedWidth =10452
                    LayoutCachedHeight =3552
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =7620
                    Top =3276
                    Width =1560
                    Height =276
                    FontSize =8
                    Name ="Label20"
                    Caption ="Coverage Recv"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =7620
                    LayoutCachedTop =3276
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =3552
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =5988
                    Top =3276
                    Width =1560
                    Height =276
                    FontSize =8
                    Name ="Label14"
                    Caption ="Coverage Required"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =5988
                    LayoutCachedTop =3276
                    LayoutCachedWidth =7548
                    LayoutCachedHeight =3552
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =11784
                    Top =3276
                    Width =1080
                    Height =276
                    FontSize =8
                    Name ="Label40"
                    Caption ="Accepted"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =11784
                    LayoutCachedTop =3276
                    LayoutCachedWidth =12864
                    LayoutCachedHeight =3552
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =11784
                    Top =3624
                    Width =360
                    Name ="EmptyCell394"
                    GroupTable =29
                    LayoutCachedLeft =11784
                    LayoutCachedTop =3624
                    LayoutCachedWidth =12144
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =11784
                    Top =3936
                    Width =360
                    Name ="EmptyCell395"
                    GroupTable =29
                    LayoutCachedLeft =11784
                    LayoutCachedTop =3936
                    LayoutCachedWidth =12144
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =12924
                    Top =3624
                    Width =192
                    Name ="EmptyCell396"
                    GroupTable =29
                    LayoutCachedLeft =12924
                    LayoutCachedTop =3624
                    LayoutCachedWidth =13116
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =12924
                    Top =3276
                    Width =804
                    Height =276
                    FontSize =8
                    Name ="Label23"
                    Caption ="Blanket"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =12924
                    LayoutCachedTop =3276
                    LayoutCachedWidth =13728
                    LayoutCachedHeight =3552
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4380
                    Top =3276
                    Width =1536
                    Height =276
                    FontSize =8
                    Name ="Label413"
                    Caption ="Type"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =4380
                    LayoutCachedTop =3276
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =3552
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =15192
                    Top =1548
                    Width =612
                    Height =252
                    Name ="EmptyCell445"
                    GroupTable =18
                    LayoutCachedLeft =15192
                    LayoutCachedTop =1548
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =15864
                    Top =1548
                    Width =996
                    Height =252
                    Name ="EmptyCell446"
                    GroupTable =18
                    LayoutCachedLeft =15864
                    LayoutCachedTop =1548
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =14052
                    Top =1548
                    Width =1080
                    Height =252
                    Name ="EmptyCell447"
                    GroupTable =18
                    LayoutCachedLeft =14052
                    LayoutCachedTop =1548
                    LayoutCachedWidth =15132
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =15192
                    Top =1872
                    Width =612
                    Height =252
                    Name ="EmptyCell448"
                    GroupTable =18
                    LayoutCachedLeft =15192
                    LayoutCachedTop =1872
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =15864
                    Top =1872
                    Width =996
                    Height =252
                    Name ="EmptyCell449"
                    GroupTable =18
                    LayoutCachedLeft =15864
                    LayoutCachedTop =1872
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =13788
                    Top =3276
                    Width =1224
                    Height =276
                    FontSize =8
                    Name ="Label450"
                    Caption ="Stip Loss Req"
                    FontName ="Segoe UI"
                    GroupTable =29
                    LayoutCachedLeft =13788
                    LayoutCachedTop =3276
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =3552
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =14052
                    Top =2520
                    Width =1752
                    Height =288
                    Name ="EmptyCell451"
                    GroupTable =18
                    LayoutCachedLeft =14052
                    LayoutCachedTop =2520
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =2808
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =15864
                    Top =2520
                    Width =996
                    Height =288
                    Name ="EmptyCell452"
                    GroupTable =18
                    LayoutCachedLeft =15864
                    LayoutCachedTop =2520
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =2808
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =13788
                    Top =3936
                    Width =1224
                    Name ="EmptyCell455"
                    GroupTable =29
                    LayoutCachedLeft =13788
                    LayoutCachedTop =3936
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =13788
                    Top =3624
                    Width =408
                    Name ="EmptyCell458"
                    GroupTable =29
                    LayoutCachedLeft =13788
                    LayoutCachedTop =3624
                    LayoutCachedWidth =14196
                    LayoutCachedHeight =3864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =10050
                    Top =165
                    Width =6870
                    Height =2955
                    Name ="Box460"
                    LayoutCachedLeft =10050
                    LayoutCachedTop =165
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =3120
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =10155
                    Top =30
                    Width =1545
                    Height =345
                    Name ="Label259"
                    Caption ="Guaranty Info"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =10155
                    LayoutCachedTop =30
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =375
                    BackThemeColorIndex =1
                End
                Begin EmptyCell
                    Left =15192
                    Top =2196
                    Width =612
                    Height =252
                    Name ="EmptyCell471"
                    GroupTable =18
                    LayoutCachedLeft =15192
                    LayoutCachedTop =2196
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =2448
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =15864
                    Top =2196
                    Width =996
                    Height =252
                    Name ="EmptyCell472"
                    GroupTable =18
                    LayoutCachedLeft =15864
                    LayoutCachedTop =2196
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =2448
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =13188
                    Top =2520
                    Width =792
                    Height =288
                    Name ="EmptyCell474"
                    GroupTable =18
                    LayoutCachedLeft =13188
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =2808
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =158
                    Width =1980
                    Height =6765
                    FontSize =9
                    Name ="MasterLeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.Client, MstrLease.MLNo, MstrLease.MLOrig FROM M"
                        "strLease WHERE (((MstrLease.Client)=[forms]![ClientsView]![clientid])) ORDER BY "
                        "MstrLease.MLOrig, MstrLease.MLNo;"
                    ColumnWidths ="0;0;1872;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =158
                    LayoutCachedWidth =2055
                    LayoutCachedHeight =6923
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =15930
                    Top =3990
                    Width =585
                    Height =210
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =41
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedLeft =15930
                    LayoutCachedTop =3990
                    LayoutCachedWidth =16515
                    LayoutCachedHeight =4200
                End
                Begin EmptyCell
                    Left =13176
                    Top =3936
                    Width =552
                    Name ="EmptyCell482"
                    GroupTable =29
                    LayoutCachedLeft =13176
                    LayoutCachedTop =3936
                    LayoutCachedWidth =13728
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin EmptyCell
                    Left =12924
                    Top =3936
                    Width =192
                    Name ="EmptyCell486"
                    GroupTable =29
                    LayoutCachedLeft =12924
                    LayoutCachedTop =3936
                    LayoutCachedWidth =13116
                    LayoutCachedHeight =4176
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =4
                    GroupTable =29
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8952
                    Top =696
                    Width =720
                    Height =252
                    FontSize =9
                    TabIndex =4
                    ForeColor =5026082
                    Name ="MLLessorID"
                    ControlSource ="MLLessorID"
                    GroupTable =12

                    LayoutCachedLeft =8952
                    LayoutCachedTop =696
                    LayoutCachedWidth =9672
                    LayoutCachedHeight =948
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =8952
                    Top =1020
                    Width =720
                    Height =252
                    Name ="EmptyCell568"
                    GroupTable =12
                    BottomPadding =150
                    LayoutCachedLeft =8952
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9672
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =8952
                    Top =372
                    Width =720
                    Height =252
                    Name ="EmptyCell569"
                    GroupTable =12
                    LayoutCachedLeft =8952
                    LayoutCachedTop =372
                    LayoutCachedWidth =9672
                    LayoutCachedHeight =624
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9576
                    Top =2124
                    Width =336
                    Height =252
                    FontSize =9
                    TabIndex =21
                    ForeColor =5026082
                    Name ="MLSignerID"
                    ControlSource ="MLSignerID"
                    GroupTable =1

                    LayoutCachedLeft =9576
                    LayoutCachedTop =2124
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =2376
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9576
                    Top =2448
                    Width =336
                    Height =252
                    FontSize =9
                    TabIndex =24
                    ForeColor =5026082
                    Name ="MLSigner2ID"
                    ControlSource ="MLSigner2ID"
                    GroupTable =1

                    LayoutCachedLeft =9576
                    LayoutCachedTop =2448
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =2700
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9576
                    Top =2772
                    Width =336
                    Height =252
                    FontSize =9
                    TabIndex =27
                    ForeColor =5026082
                    Name ="MLSigner3ID"
                    ControlSource ="MLSigner3ID"
                    GroupTable =1

                    LayoutCachedLeft =9576
                    LayoutCachedTop =2772
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =3024
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =9576
                    Top =1800
                    Width =336
                    Height =255
                    Name ="EmptyCell573"
                    GroupTable =1
                    LayoutCachedLeft =9576
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =2055
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16932
                    Top =600
                    Width =360
                    Height =225
                    FontSize =9
                    TabIndex =9
                    ForeColor =5026082
                    Name ="MLGuarSignerID"
                    ControlSource ="MLGuarSignerID"
                    GroupTable =18

                    LayoutCachedLeft =16932
                    LayoutCachedTop =600
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =825
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =900
                    Width =360
                    Height =252
                    Name ="EmptyCell579"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =900
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =1152
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =1224
                    Width =360
                    Height =252
                    Name ="EmptyCell580"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =1224
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =1476
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =1548
                    Width =360
                    Height =252
                    Name ="EmptyCell581"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =1548
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =1800
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =1872
                    Width =360
                    Height =252
                    Name ="EmptyCell582"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =1872
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =2196
                    Width =360
                    Height =252
                    Name ="EmptyCell583"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =2196
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =2448
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =16932
                    Top =2520
                    Width =360
                    Height =288
                    Name ="EmptyCell584"
                    GroupTable =18
                    LayoutCachedLeft =16932
                    LayoutCachedTop =2520
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =2808
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =18
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =3420
                    Top =2124
                    Width =3084
                    Height =252
                    FontSize =8
                    TabIndex =19
                    Name ="MLSigner"
                    ControlSource ="MLSigner"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3488 Order by FullName;"
                    ColumnWidths ="2880;1440;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2124
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =2376
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2472
                            Top =2124
                            Width =876
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label81"
                            Caption ="1st Signer"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =2472
                            LayoutCachedTop =2124
                            LayoutCachedWidth =3348
                            LayoutCachedHeight =2376
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
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =3420
                    Top =2448
                    Width =3084
                    Height =252
                    FontSize =8
                    TabIndex =22
                    Name ="MLSigner2"
                    ControlSource ="MLSigner2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3488 Order by FullName;"
                    ColumnWidths ="2880;1440;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2448
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =2700
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2472
                            Top =2448
                            Width =876
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label89"
                            Caption ="2nd Signer"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =2472
                            LayoutCachedTop =2448
                            LayoutCachedWidth =3348
                            LayoutCachedHeight =2700
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =3420
                    Top =2772
                    Width =3084
                    Height =252
                    FontSize =8
                    TabIndex =25
                    Name ="MLSigner3"
                    ControlSource ="MLSigner3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3488 Order by FullName;"
                    ColumnWidths ="2880;1440;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2772
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =3024
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2472
                            Top =2772
                            Width =876
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label87"
                            Caption ="3rd Signer"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =2472
                            LayoutCachedTop =2772
                            LayoutCachedWidth =3348
                            LayoutCachedHeight =3024
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5760
                    Left =12048
                    Top =900
                    Width =4812
                    Height =252
                    FontSize =8
                    TabIndex =10
                    Name ="MLGuarSigner"
                    ControlSource ="MLGuarSigner"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FullName, Title, contactID, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3855 Order by FullName;"
                    ColumnWidths ="2880;2880;0"
                    StatusBarText ="link to contacts table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =18

                    LayoutCachedLeft =12048
                    LayoutCachedTop =900
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =1152
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =10368
                            Top =900
                            Width =1608
                            Height =252
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label261"
                            Caption ="Signer's Name"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =10368
                            LayoutCachedTop =900
                            LayoutCachedWidth =11976
                            LayoutCachedHeight =1152
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
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
' See "Leasing_sub.cls"
