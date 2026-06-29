Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3945
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =14055
    Top =8730
    Right =17910
    Bottom =10560
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf24424eab747e440
    End
    RecordSource ="SELECT MstrLease.MLNo, Schedule.SchID, Schedule.SchNo, Schedule.UpFrontComm, Sch"
        "edule.FundCommPmt, Schedule.SchIntCommbase FROM MstrLease INNER JOIN Schedule ON"
        " MstrLease.MLID = Schedule.SchMLRef WHERE (((Schedule.SchID)=[forms]![AcctMenu]!"
        "[SelSch]));"
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
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =1566
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =510
                    Top =195
                    Width =1379
                    Height =209
                    FontSize =9
                    FontWeight =700
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    Format =">"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =510
                    LayoutCachedTop =195
                    LayoutCachedWidth =1889
                    LayoutCachedHeight =404
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1995
                    Top =195
                    Width =1560
                    Height =209
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    Format =">"
                    StatusBarText ="schedule number starts with 1"

                    LayoutCachedLeft =1995
                    LayoutCachedTop =195
                    LayoutCachedWidth =3555
                    LayoutCachedHeight =404
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =840
                    Width =1500
                    Height =312
                    FontSize =10
                    TabIndex =2
                    BorderColor =12566463
                    Name ="UpFrontComm"
                    ControlSource ="UpFrontComm"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="commission paid upfront"
                    ValidationRule =">=0"
                    ValidationText ="Commisin paid upfron must be >=0"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =1980
                    LayoutCachedTop =840
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1152
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
                            OverlapFlags =85
                            TextAlign =1
                            Left =528
                            Top =840
                            Width =1380
                            Height =312
                            FontSize =10
                            Name ="Label2"
                            Caption ="Up Front:"
                            GroupTable =1
                            LayoutCachedLeft =528
                            LayoutCachedTop =840
                            LayoutCachedWidth =1908
                            LayoutCachedHeight =1152
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =1224
                    Width =1500
                    Height =312
                    FontSize =10
                    TabIndex =3
                    BorderColor =12566463
                    Name ="FundCommPmt"
                    ControlSource ="FundCommPmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amount of comomission paid at funding with prepayment adjustment."
                    ValidationRule =">=0"
                    ValidationText ="Commision at funding must be .=0"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1224
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1536
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
                            OverlapFlags =85
                            TextAlign =1
                            Left =528
                            Top =1224
                            Width =1380
                            Height =312
                            FontSize =10
                            Name ="Label3"
                            Caption ="@ Funding:"
                            GroupTable =1
                            LayoutCachedLeft =528
                            LayoutCachedTop =1224
                            LayoutCachedWidth =1908
                            LayoutCachedHeight =1536
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =456
                    Width =1500
                    Height =312
                    FontSize =10
                    TabIndex =4
                    BorderColor =12566463
                    Name ="SchIntCommbase"
                    ControlSource ="SchIntCommbase"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="basis for intial commission calculationa"
                    GroupTable =1

                    LayoutCachedLeft =1980
                    LayoutCachedTop =456
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =768
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =528
                            Top =456
                            Width =1380
                            Height =312
                            FontSize =10
                            Name ="Label5"
                            Caption ="Esti Comm %:"
                            GroupTable =1
                            LayoutCachedLeft =528
                            LayoutCachedTop =456
                            LayoutCachedWidth =1908
                            LayoutCachedHeight =768
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =510
                    Top =15
                    Width =2370
                    Height =195
                    FontSize =8
                    FontWeight =700
                    Name ="Label6"
                    Caption ="MasterLease                | Schdule #"
                    LayoutCachedLeft =510
                    LayoutCachedTop =15
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =210
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
            End
        End
    End
End
