Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
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
    Width =3960
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =14265
    Top =7620
    Right =18135
    Bottom =8580
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x78f6820c5020e540
    End
    RecordSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName, DealCommissions.DealComm"
        "ission, DealCommissions.dealOutcomeDate, vw_SixKeys.SchID FROM vw_SixKeys INNER "
        "JOIN DealCommissions ON (vw_SixKeys.FAID = DealCommissions.FAID) AND (vw_SixKeys"
        ".clientID = DealCommissions.clientID);"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin Section
            Height =1215
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1545
                    Top =480
                    Width =1365
                    Height =315
                    Name ="ClientComm%"
                    ControlSource ="DealCommission"
                    Format ="Percent"
                    StatusBarText ="New deal percentage effective 1 year from ClientCommDt"
                    EventProcPrefix ="ClientComm_"

                    LayoutCachedLeft =1545
                    LayoutCachedTop =480
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =795
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =540
                            Width =1080
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label2"
                            Caption ="5. Comm %:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =540
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1545
                    Top =900
                    Width =1350
                    Height =315
                    TabIndex =1
                    Name ="ClientCommdt"
                    ControlSource ="dealOutcomeDate"
                    Format ="mm/dd/yyyy"

                    LayoutCachedLeft =1545
                    LayoutCachedTop =900
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =900
                            Width =1050
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label3"
                            Caption ="6. Awarded"
                            LayoutCachedLeft =60
                            LayoutCachedTop =900
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =1185
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3165
                    Top =525
                    Width =510
                    Height =315
                    FontSize =9
                    TabIndex =2
                    Name ="ClientID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =3165
                    LayoutCachedTop =525
                    LayoutCachedWidth =3675
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
            End
        End
    End
End
