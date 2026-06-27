Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15336
    DatasheetFontHeight =11
    ItemSuffix =78
    Left =4215
    Top =6840
    Right =19800
    Bottom =14625
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xcb7dd40b9021e540
    End
    RecordSource ="SELECT MstrLease.*, Clients.clientTaxID FROM MstrLease LEFT JOIN Clients ON Mstr"
        "Lease.MLGuarantor = Clients.clientID ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
    OnCurrent ="[Event Procedure]"
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
            Height =480
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    Width =2040
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label40"
                    Caption ="Master Lease"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =300
                    Width =15336
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =300
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =480
                    TabIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =13995
                    Top =30
                    Width =840
                    Height =300
                    FontSize =8
                    Name ="close"
                    Caption ="Close "
                    FontName ="Segoe UI"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"close\" xmlns=\"http://schemas.microsoft.com/office/accessse"
                                "rvices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Stateme"
                                "nts></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =13995
                    LayoutCachedTop =30
                    LayoutCachedWidth =14835
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =8872517
                    HoverColor =8872517
                    PressedColor =15527148
                    PressedForeThemeColorIndex =0
                    QuickStyle =2
                    QuickStyleMask =-241
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7200
            Name ="Detail"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =75
                    Top =135
                    Width =7830
                    Height =1200
                    Name ="Box476"
                    LayoutCachedLeft =75
                    LayoutCachedTop =135
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =1335
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =75
                    Top =1575
                    Width =7830
                    Height =1530
                    Name ="Box475"
                    LayoutCachedLeft =75
                    LayoutCachedTop =1575
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =3105
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =75
                    Top =3390
                    Width =14775
                    Height =900
                    Name ="Box415"
                    LayoutCachedLeft =75
                    LayoutCachedTop =3390
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =4290
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =975
                    Width =1200
                    Height =255
                    FontSize =8
                    TabIndex =4
                    Name ="LeaseDate"
                    ControlSource ="MLDT"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1320
                    LayoutCachedTop =975
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =390
                            Top =975
                            Width =870
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label253"
                            Caption ="Lease Date"
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =390
                            LayoutCachedTop =975
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4485
                    Top =2085
                    Width =2925
                    FontSize =8
                    TabIndex =19
                    Name ="MLSignerTitle1"
                    ControlSource ="=[MLSigner].[Column](2)"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =4485
                    LayoutCachedTop =2085
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =2325
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =1155
                    Width =4800
                    Height =255
                    FontSize =8
                    TabIndex =10
                    Name ="GuarantorTitle"
                    ControlSource ="=[GuarantorSigner].[column](2)"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =1155
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =1410
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =1155
                            Width =1605
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label262"
                            Caption ="Signer's Title"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =1155
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =1410
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =1470
                    Width =1935
                    Height =255
                    FontSize =8
                    TabIndex =11
                    Name ="LeaseGtyTaxID"
                    ControlSource ="clienttaxid"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =1470
                    LayoutCachedWidth =11895
                    LayoutCachedHeight =1725
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =1470
                            Width =1605
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label263"
                            Caption ="Guarantor Tax ID"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =1470
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =1725
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4485
                    Top =2385
                    Width =2925
                    FontSize =8
                    TabIndex =22
                    Name ="MLSignerTitle2"
                    ControlSource ="=[MLSigner2].[Column](2)"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =4485
                    LayoutCachedTop =2385
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =2625
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4485
                    Top =2685
                    Width =2925
                    FontSize =8
                    TabIndex =25
                    Name ="MLSignerTitle3"
                    ControlSource ="=[MLSigner3].[Column](2)"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =4485
                    LayoutCachedTop =2685
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =2925
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3405
                    Top =975
                    Width =1200
                    Height =255
                    FontSize =8
                    TabIndex =5
                    Name ="MLSent"
                    ControlSource ="MLSent"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =3405
                    LayoutCachedTop =975
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2580
                            Top =975
                            Width =765
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label11"
                            Caption ="Sent"
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =2580
                            LayoutCachedTop =975
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5550
                    Top =975
                    Width =1200
                    Height =255
                    FontSize =8
                    TabIndex =6
                    Name ="MLBack"
                    ControlSource ="MLBack"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =5550
                    LayoutCachedTop =975
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4665
                            Top =975
                            Width =825
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label12"
                            Caption ="Returned"
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =4665
                            LayoutCachedTop =975
                            LayoutCachedWidth =5490
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =210
                    Top =3180
                    Width =990
                    Height =345
                    Name ="Label13"
                    Caption ="Insurance"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedTop =3180
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3525
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3885
                    Top =3570
                    Width =1560
                    FontSize =8
                    TabIndex =28
                    BorderColor =8872517
                    Name ="MLINReq"
                    ControlSource ="MLINReq"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Minimum Insurance Amount Required"
                    ValidationRule =">=0"
                    ValidationText ="Liability Amt Req must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =3885
                    LayoutCachedTop =3570
                    LayoutCachedWidth =5445
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =12120
                    Top =3570
                    Width =750
                    TabIndex =34
                    Name ="MLStipLoss"
                    ControlSource ="MLStipLoss"
                    StatusBarText ="is a Stip Loss table require in each lease"
                    GroupTable =4

                    LayoutCachedLeft =12120
                    LayoutCachedTop =3570
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5505
                    Top =3570
                    Width =1560
                    FontSize =8
                    TabIndex =29
                    BorderColor =8872517
                    Name ="MLINRecA"
                    ControlSource ="MLINRecA"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="insurance amount received"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =5505
                    LayoutCachedTop =3570
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7125
                    Top =3570
                    Width =1200
                    FontSize =8
                    TabIndex =30
                    BorderColor =8872517
                    Name ="MLInsRec"
                    ControlSource ="MLInsRec"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =7125
                    LayoutCachedTop =3570
                    LayoutCachedWidth =8325
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =11040
                    Top =3570
                    Width =555
                    TabIndex =33
                    Name ="MLINSBLANKET"
                    ControlSource ="MLINSBLANKET"
                    StatusBarText ="IF YES AT LEASE LEVEL IS NO AT SCHEDULE LEVEL"
                    GroupTable =4

                    LayoutCachedLeft =11040
                    LayoutCachedTop =3570
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2295
                    Top =3570
                    Width =1530
                    FontSize =8
                    TabIndex =27
                    BorderColor =8872517
                    Name ="MLPhysIN"
                    ControlSource ="MLLiabIns"
                    RowSourceType ="Value List"
                    RowSource ="S;self-insured;R;Required"
                    ColumnWidths ="0;1440"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =2295
                    LayoutCachedTop =3570
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =1365
                            Top =3570
                            Width =870
                            Height =240
                            FontSize =8
                            Name ="Label24"
                            Caption ="Liability"
                            FontName ="Segoe UI"
                            GroupTable =4
                            LayoutCachedLeft =1365
                            LayoutCachedTop =3570
                            LayoutCachedWidth =2235
                            LayoutCachedHeight =3810
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
                            BackThemeColorIndex =8
                            BackTint =40.0
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9960
                    Top =570
                    Width =4800
                    Height =225
                    FontSize =8
                    TabIndex =7
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
                    GroupTable =3
                    AllowValueListEdits =0

                    LayoutCachedLeft =9960
                    LayoutCachedTop =570
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =795
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =570
                            Width =1605
                            Height =225
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label25"
                            Caption ="Guarantor"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =570
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =795
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =11955
                    Top =1785
                    Width =1080
                    Height =255
                    TabIndex =13
                    Name ="MLGuarUnlim"
                    ControlSource ="MLGuarUnlim"
                    StatusBarText ="Unlimited Guaranty?"
                    GroupTable =3

                    LayoutCachedLeft =11955
                    LayoutCachedTop =1785
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =11100
                            Top =1785
                            Width =795
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label26"
                            Caption ="Unlimited"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =11100
                            LayoutCachedTop =1785
                            LayoutCachedWidth =11895
                            LayoutCachedHeight =2040
                            RowStart =4
                            RowEnd =4
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =1785
                    Width =1080
                    Height =255
                    FontSize =8
                    TabIndex =12
                    Name ="MLGuarLimit"
                    ControlSource ="MLGuarLimit"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText =" Limit must by >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =1785
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =1785
                            Width =1605
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label27"
                            Caption ="Guaranty Limit"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =1785
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =2040
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =2415
                    Width =1080
                    Height =285
                    FontSize =8
                    TabIndex =16
                    Name ="MLGuarExp"
                    ControlSource ="MLGuarExp"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =2415
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2700
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =2415
                            Width =1605
                            Height =285
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label28"
                            Caption ="Guaranty Expiration"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =2415
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =2700
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =2100
                    Width =1080
                    Height =255
                    FontSize =8
                    TabIndex =14
                    Name ="MLGuarSent"
                    ControlSource ="MLGuarSent"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =2100
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2355
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8295
                            Top =2100
                            Width =1605
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label29"
                            Caption ="Guaranty Sent"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =2100
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =2355
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11955
                    Top =2100
                    Width =1080
                    Height =255
                    FontSize =8
                    TabIndex =15
                    Name ="MLGuarBack"
                    ControlSource ="MLGuarBack"
                    FontName ="Segoe UI"
                    GroupTable =3

                    LayoutCachedLeft =11955
                    LayoutCachedTop =2100
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =2355
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =11100
                            Top =2100
                            Width =795
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label30"
                            Caption ="Returned"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =11100
                            LayoutCachedTop =2100
                            LayoutCachedWidth =11895
                            LayoutCachedHeight =2355
                            RowStart =5
                            RowEnd =5
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =4110
                    Top =1770
                    Width =315
                    Height =255
                    TabIndex =17
                    Name ="MLAnyCoASign"
                    ControlSource ="MLAnyCoASign"
                    StatusBarText ="if No signer must be incumbent"
                    GroupTable =2

                    LayoutCachedLeft =4110
                    LayoutCachedTop =1770
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =2025
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2370
                            Top =1770
                            Width =1680
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label31"
                            Caption ="Any Signer-CofA"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =2370
                            LayoutCachedTop =1770
                            LayoutCachedWidth =4050
                            LayoutCachedHeight =2025
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =210
                    Width =1080
                    Height =345
                    Name ="Label258"
                    Caption ="Lease Info"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =345
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1320
                    Top =345
                    Width =2670
                    Height =255
                    FontSize =8
                    BackColor =11271167
                    BorderColor =8872517
                    Name ="Text251"
                    ControlSource ="MLno"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =1320
                    LayoutCachedTop =345
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =600
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =390
                            Top =345
                            Width =870
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label252"
                            Caption ="MLA No."
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =390
                            LayoutCachedTop =345
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =600
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
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
                    Left =4665
                    Top =345
                    Width =2085
                    Height =255
                    FontSize =8
                    TabIndex =1
                    BackColor =11271167
                    BorderColor =8872517
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    RowSourceType ="Value List"
                    RowSource ="FA;FA Lease;Other;Other Leasing Entity;FA Rmkt;Purchased for Remarketing;Portfol"
                        "io;Tracked Units"
                    ColumnWidths ="1080;3240"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="\"FA\""
                    FontName ="Segoe UI"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =1
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4665
                    LayoutCachedTop =345
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =600
                    ColumnStart =5
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =4050
                            Top =345
                            Width =555
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label2"
                            Caption ="Orig:"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =4050
                            LayoutCachedTop =345
                            LayoutCachedWidth =4605
                            LayoutCachedHeight =600
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =10065
                    Top =3570
                    Width =660
                    TabIndex =32
                    Name ="MLINSAccpt"
                    ControlSource ="MLINSAccpt"
                    GroupTable =4

                    LayoutCachedLeft =10065
                    LayoutCachedTop =3570
                    LayoutCachedWidth =10725
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8385
                    Top =3570
                    Width =1200
                    FontSize =8
                    TabIndex =31
                    BorderColor =8872517
                    Name ="Text41"
                    ControlSource ="MLInsExp"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =8385
                    LayoutCachedTop =3570
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7125
                    Top =3870
                    Width =1200
                    FontSize =8
                    TabIndex =36
                    BorderColor =8872517
                    Name ="MLPhyRec"
                    ControlSource ="MLPhyRec"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =7125
                    LayoutCachedTop =3870
                    LayoutCachedWidth =8325
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1320
                    Top =660
                    Width =5430
                    Height =255
                    FontSize =8
                    TabIndex =2
                    Name ="MLLessor"
                    ControlSource ="MLLessorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.clientID, Clients.ClientShNm FROM Clients WHERE (((Clients.Client"
                        "ShNm)<>\"lessor1\" And (Clients.ClientShNm)<>\"lessor2\") AND ((Clients.clientTy"
                        "pe)=\"bank\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="0;4320"
                    StatusBarText ="If other name of lessor"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =1320
                    LayoutCachedTop =660
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =390
                            Top =660
                            Width =870
                            Height =255
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label3"
                            Caption ="Lessor"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =390
                            LayoutCachedTop =660
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =915
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
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
                    Left =2295
                    Top =3870
                    Width =1530
                    FontSize =8
                    TabIndex =35
                    BorderColor =8872517
                    Name ="Combo54"
                    ControlSource ="MLPhysIN"
                    RowSourceType ="Value List"
                    RowSource ="S;self-insured;R;Required"
                    ColumnWidths ="0;1440"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =2295
                    LayoutCachedTop =3870
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =1365
                            Top =3870
                            Width =870
                            Height =240
                            FontSize =8
                            Name ="Label55"
                            Caption ="Physical"
                            FontName ="Segoe UI"
                            GroupTable =4
                            LayoutCachedLeft =1365
                            LayoutCachedTop =3870
                            LayoutCachedWidth =2235
                            LayoutCachedHeight =4110
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            BackThemeColorIndex =9
                            BackTint =40.0
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =10065
                    Top =3870
                    Width =660
                    TabIndex =38
                    Name ="PhysIncAccpt"
                    ControlSource ="PhysIncAccpt"
                    GroupTable =4

                    LayoutCachedLeft =10065
                    LayoutCachedTop =3870
                    LayoutCachedWidth =10725
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2295
                    Top =4170
                    Width =10575
                    FontSize =8
                    TabIndex =39
                    BorderColor =10040879
                    Name ="INScomment"
                    ControlSource ="INScomment"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =2295
                    LayoutCachedTop =4170
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =4410
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =11
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =1365
                            Top =4170
                            Width =870
                            Height =240
                            FontSize =8
                            Name ="Label57"
                            Caption ="Comment"
                            FontName ="Segoe UI"
                            GroupTable =4
                            LayoutCachedLeft =1365
                            LayoutCachedTop =4170
                            LayoutCachedWidth =2235
                            LayoutCachedHeight =4410
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =4
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8385
                    Top =3870
                    Width =1200
                    FontSize =8
                    TabIndex =37
                    BorderColor =8872517
                    Name ="PhysInsExp"
                    ControlSource ="PhysInsExp"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =4

                    LayoutCachedLeft =8385
                    LayoutCachedTop =3870
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Top =4440
                    Width =15285
                    Height =2760
                    TabIndex =42
                    BorderColor =12632256
                    Name ="MLAmendf"
                    SourceObject ="Form.MLAmendf"
                    LinkChildFields ="MLRef"
                    LinkMasterFields ="MLID"
                    GridlineColor =0

                    LayoutCachedTop =4440
                    LayoutCachedWidth =15285
                    LayoutCachedHeight =7200
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BorderShade =100.0
                    ShowPageHeaderAndPageFooter =0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4485
                    Top =1770
                    Width =2925
                    Height =255
                    FontSize =8
                    ForeColor =10040879
                    Name ="Label4"
                    Caption ="Title"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =4485
                    LayoutCachedTop =1770
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =2025
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =1335
                    Top =1770
                    Width =975
                    Height =255
                    FontSize =8
                    ForeColor =10040879
                    Name ="Label91"
                    Caption ="Full Name"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =1335
                    LayoutCachedTop =1770
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =2025
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =405
                    Top =1770
                    Width =870
                    Height =255
                    Name ="EmptyCell94"
                    GroupTable =2
                    LayoutCachedLeft =405
                    LayoutCachedTop =1770
                    LayoutCachedWidth =1275
                    LayoutCachedHeight =2025
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =210
                    Top =1380
                    Width =765
                    Height =345
                    Name ="Label209"
                    Caption ="Signers"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedTop =1380
                    LayoutCachedWidth =975
                    LayoutCachedHeight =1725
                    BackThemeColorIndex =1
                End
                Begin EmptyCell
                    Left =3885
                    Top =3870
                    Width =1560
                    Name ="EmptyCell331"
                    GroupTable =4
                    LayoutCachedLeft =3885
                    LayoutCachedTop =3870
                    LayoutCachedWidth =5445
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =8385
                    Top =3240
                    Width =1200
                    Height =270
                    FontSize =8
                    Name ="Label356"
                    Caption ="Expires "
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =8385
                    LayoutCachedTop =3240
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =3510
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =1365
                    Top =3240
                    Width =870
                    Height =270
                    Name ="EmptyCell359"
                    GroupTable =4
                    LayoutCachedLeft =1365
                    LayoutCachedTop =3240
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =3510
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =7125
                    Top =3240
                    Width =1200
                    Height =270
                    FontSize =8
                    Name ="Label21"
                    Caption ="Date Recv"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =7125
                    LayoutCachedTop =3240
                    LayoutCachedWidth =8325
                    LayoutCachedHeight =3510
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =5505
                    Top =3240
                    Width =1560
                    Height =270
                    FontSize =8
                    Name ="Label20"
                    Caption ="Coverage Recv"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =5505
                    LayoutCachedTop =3240
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =3510
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3885
                    Top =3240
                    Width =1560
                    Height =270
                    FontSize =8
                    Name ="Label14"
                    Caption ="Coverage Required"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =3885
                    LayoutCachedTop =3240
                    LayoutCachedWidth =5445
                    LayoutCachedHeight =3510
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =9645
                    Top =3240
                    Width =1080
                    Height =270
                    FontSize =8
                    Name ="Label45"
                    Caption ="Accepted"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =9645
                    LayoutCachedTop =3240
                    LayoutCachedWidth =10725
                    LayoutCachedHeight =3510
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9645
                    Top =3570
                    Width =360
                    Name ="EmptyCell394"
                    GroupTable =4
                    LayoutCachedLeft =9645
                    LayoutCachedTop =3570
                    LayoutCachedWidth =10005
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9645
                    Top =3870
                    Width =360
                    Name ="EmptyCell395"
                    GroupTable =4
                    LayoutCachedLeft =9645
                    LayoutCachedTop =3870
                    LayoutCachedWidth =10005
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =10785
                    Top =3570
                    Width =195
                    Name ="EmptyCell396"
                    GroupTable =4
                    LayoutCachedLeft =10785
                    LayoutCachedTop =3570
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =10785
                    Top =3240
                    Width =810
                    Height =270
                    FontSize =8
                    Name ="Label23"
                    Caption ="Blanket"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =10785
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =3510
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =2295
                    Top =3240
                    Width =1530
                    Height =270
                    FontSize =8
                    Name ="Label413"
                    Caption ="Type"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =2295
                    LayoutCachedTop =3240
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =3510
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =13095
                    Top =1470
                    Width =615
                    Height =255
                    Name ="EmptyCell445"
                    GroupTable =3
                    LayoutCachedLeft =13095
                    LayoutCachedTop =1470
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =1725
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =13770
                    Top =1470
                    Width =990
                    Height =255
                    Name ="EmptyCell446"
                    GroupTable =3
                    LayoutCachedLeft =13770
                    LayoutCachedTop =1470
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =1725
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11955
                    Top =1470
                    Width =1080
                    Height =255
                    Name ="EmptyCell447"
                    GroupTable =3
                    LayoutCachedLeft =11955
                    LayoutCachedTop =1470
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =1725
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =13095
                    Top =1785
                    Width =615
                    Height =255
                    Name ="EmptyCell448"
                    GroupTable =3
                    LayoutCachedLeft =13095
                    LayoutCachedTop =1785
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =13770
                    Top =1785
                    Width =990
                    Height =255
                    Name ="EmptyCell449"
                    GroupTable =3
                    LayoutCachedLeft =13770
                    LayoutCachedTop =1785
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =11655
                    Top =3240
                    Width =1215
                    Height =270
                    FontSize =8
                    Name ="Label450"
                    Caption ="Stip Loss Req"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LayoutCachedLeft =11655
                    LayoutCachedTop =3240
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =3510
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =11955
                    Top =2415
                    Width =1755
                    Height =285
                    Name ="EmptyCell451"
                    GroupTable =3
                    LayoutCachedLeft =11955
                    LayoutCachedTop =2415
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =2700
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =13770
                    Top =2415
                    Width =990
                    Height =285
                    Name ="EmptyCell452"
                    GroupTable =3
                    LayoutCachedLeft =13770
                    LayoutCachedTop =2415
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =2700
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11655
                    Top =3870
                    Width =1215
                    Name ="EmptyCell455"
                    GroupTable =4
                    LayoutCachedLeft =11655
                    LayoutCachedTop =3870
                    LayoutCachedWidth =12870
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =11655
                    Top =3570
                    Width =405
                    Name ="EmptyCell458"
                    GroupTable =4
                    LayoutCachedLeft =11655
                    LayoutCachedTop =3570
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =3810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =7980
                    Top =135
                    Width =6870
                    Height =2955
                    Name ="Box460"
                    LayoutCachedLeft =7980
                    LayoutCachedTop =135
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =3090
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =8085
                    Width =1545
                    Height =345
                    Name ="Label259"
                    Caption ="Guaranty Info"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =8085
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =345
                    BackThemeColorIndex =1
                End
                Begin EmptyCell
                    Left =13095
                    Top =2100
                    Width =615
                    Height =255
                    Name ="EmptyCell471"
                    GroupTable =3
                    LayoutCachedLeft =13095
                    LayoutCachedTop =2100
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =2355
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =13770
                    Top =2100
                    Width =990
                    Height =255
                    Name ="EmptyCell472"
                    GroupTable =3
                    LayoutCachedLeft =13770
                    LayoutCachedTop =2100
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =2355
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11100
                    Top =2415
                    Width =795
                    Height =285
                    Name ="EmptyCell474"
                    GroupTable =3
                    LayoutCachedLeft =11100
                    LayoutCachedTop =2415
                    LayoutCachedWidth =11895
                    LayoutCachedHeight =2700
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13860
                    Top =3960
                    Width =585
                    Height =210
                    FontSize =9
                    TabIndex =41
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedLeft =13860
                    LayoutCachedTop =3960
                    LayoutCachedWidth =14445
                    LayoutCachedHeight =4170
                End
                Begin EmptyCell
                    Left =11040
                    Top =3870
                    Width =555
                    Name ="EmptyCell482"
                    GroupTable =4
                    LayoutCachedLeft =11040
                    LayoutCachedTop =3870
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =10785
                    Top =3870
                    Width =195
                    Name ="EmptyCell486"
                    GroupTable =4
                    LayoutCachedLeft =10785
                    LayoutCachedTop =3870
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =4
                    GroupTable =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13200
                    Top =3960
                    Width =585
                    Height =210
                    FontSize =9
                    TabIndex =40
                    ForeColor =5026082
                    Name ="Client"
                    ControlSource ="Client"

                    LayoutCachedLeft =13200
                    LayoutCachedTop =3960
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =4170
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1335
                    Top =2085
                    Width =3090
                    FontSize =8
                    TabIndex =18
                    Name ="MLSigner"
                    ControlSource ="MLSignerID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1335
                    LayoutCachedTop =2085
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =2325
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =405
                            Top =2085
                            Width =870
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label81"
                            Caption ="1st Signer"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =405
                            LayoutCachedTop =2085
                            LayoutCachedWidth =1275
                            LayoutCachedHeight =2325
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =7470
                    Top =1770
                    Width =285
                    Height =255
                    Name ="EmptyCell55"
                    GroupTable =2
                    LayoutCachedLeft =7470
                    LayoutCachedTop =1770
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =2025
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7470
                    Top =2085
                    Width =285
                    FontSize =9
                    TabIndex =20
                    ForeColor =5026082
                    Name ="MLSignerID"
                    ControlSource ="MLSignerID"
                    GroupTable =2

                    LayoutCachedLeft =7470
                    LayoutCachedTop =2085
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =2325
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7470
                    Top =2385
                    Width =285
                    FontSize =9
                    TabIndex =23
                    ForeColor =5026082
                    Name ="MLSigner2ID"
                    ControlSource ="MLSigner2ID"
                    GroupTable =2

                    LayoutCachedLeft =7470
                    LayoutCachedTop =2385
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =2625
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7470
                    Top =2685
                    Width =285
                    FontSize =9
                    TabIndex =26
                    ForeColor =5026082
                    Name ="MLSigner3ID"
                    ControlSource ="MLSigner3ID"
                    GroupTable =2

                    LayoutCachedLeft =7470
                    LayoutCachedTop =2685
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =2925
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =14820
                    Top =570
                    Width =225
                    Height =225
                    Name ="EmptyCell65"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =570
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =795
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =14820
                    Top =1155
                    Width =225
                    Height =255
                    Name ="EmptyCell67"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =1155
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =1410
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =14820
                    Top =1470
                    Width =225
                    Height =255
                    Name ="EmptyCell68"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =1470
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =1725
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =14820
                    Top =1785
                    Width =225
                    Height =255
                    Name ="EmptyCell69"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =1785
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =14820
                    Top =2100
                    Width =225
                    Height =255
                    Name ="EmptyCell70"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =2100
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =2355
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =14820
                    Top =2415
                    Width =225
                    Height =285
                    Name ="EmptyCell71"
                    GroupTable =3
                    LayoutCachedLeft =14820
                    LayoutCachedTop =2415
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =2700
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14820
                    Top =855
                    Width =225
                    FontSize =9
                    TabIndex =9
                    ForeColor =5026082
                    Name ="MLGuarSignerID"
                    ControlSource ="MLGuarSignerID"
                    GroupTable =3

                    LayoutCachedLeft =14820
                    LayoutCachedTop =855
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =1095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =6810
                    Top =345
                    Width =360
                    Height =255
                    Name ="EmptyCell74"
                    GroupTable =1
                    LayoutCachedLeft =6810
                    LayoutCachedTop =345
                    LayoutCachedWidth =7170
                    LayoutCachedHeight =600
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6810
                    Top =975
                    Width =360
                    Height =255
                    Name ="EmptyCell76"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =6810
                    LayoutCachedTop =975
                    LayoutCachedWidth =7170
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6810
                    Top =660
                    Width =360
                    Height =255
                    FontSize =9
                    TabIndex =3
                    ForeColor =5026082
                    Name ="MLLessorID"
                    ControlSource ="MLLessorID"
                    GroupTable =1

                    LayoutCachedLeft =6810
                    LayoutCachedTop =660
                    LayoutCachedWidth =7170
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1335
                    Top =2385
                    Width =3090
                    FontSize =8
                    TabIndex =21
                    Name ="MLSigner2"
                    ControlSource ="MLSigner2ID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1335
                    LayoutCachedTop =2385
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =2625
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =405
                            Top =2385
                            Width =870
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label89"
                            Caption ="2nd Signer"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =405
                            LayoutCachedTop =2385
                            LayoutCachedWidth =1275
                            LayoutCachedHeight =2625
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
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
                    Left =1335
                    Top =2685
                    Width =3090
                    FontSize =8
                    TabIndex =24
                    Name ="MLSigner3"
                    ControlSource ="MLSigner3ID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1335
                    LayoutCachedTop =2685
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =2925
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =405
                            Top =2685
                            Width =870
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label87"
                            Caption ="3rd Signer"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =405
                            LayoutCachedTop =2685
                            LayoutCachedWidth =1275
                            LayoutCachedHeight =2925
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9960
                    Top =855
                    Width =4800
                    FontSize =8
                    TabIndex =8
                    Name ="GuarantorSigner"
                    ControlSource ="MLGuarSignerID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2880;2880"
                    StatusBarText ="link to contacts table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =3

                    LayoutCachedLeft =9960
                    LayoutCachedTop =855
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =1095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =8295
                            Top =855
                            Width =1605
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label261"
                            Caption ="Signer's Name"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LayoutCachedLeft =8295
                            LayoutCachedTop =855
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =1095
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin EmptyCell
                    Left =5505
                    Top =3870
                    Width =1560
                    Name ="EmptyCell77"
                    GroupTable =4
                    LayoutCachedLeft =5505
                    LayoutCachedTop =3870
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =4110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =4
                End
            End
        End
        Begin FormFooter
            Height =120
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "MstrLease.cls"
