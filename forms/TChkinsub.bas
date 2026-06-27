Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9390
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =4350
    Top =5040
    Right =13620
    Bottom =6780
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x693115c91c07e540
    End
    RecordSource ="SELECT Units.UnitVIN, Units.UnitYr AS DBUnitYR, Locations.LState AS DBTitlest, S"
        "chedule.SchTitleToTrust AS DBTitleTrust FROM (Schedule INNER JOIN SchGrp ON Sche"
        "dule.SchID = SchGrp.SGrpSchID) INNER JOIN (Units INNER JOIN Locations ON Units.U"
        "nitTitleLocationRef = Locations.LocationID) ON SchGrp.SGrpID = Units.GroupREF WH"
        "ERE (((Units.UnitVIN)=[forms]![titleCheckin]![seekvin]));"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OnLoadEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="MoveSize"
            Argument ="2880"
            Argument ="5040"
            Argument ="9360"
            Argument ="2160"
        End
    End
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
        Begin FormHeader
            Height =435
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =75
                    Width =4725
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label7"
                    Caption ="Database Information for VIN on Unit Record"
                    LayoutCachedLeft =75
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            Height =975
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3525
                    Height =315
                    Name ="DBUnitYR"
                    ControlSource ="DBUnitYR"

                    LayoutCachedLeft =3525
                    LayoutCachedWidth =4965
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2280
                            Top =15
                            Width =1140
                            Height =315
                            Name ="Label1"
                            Caption ="DB Unit YR: "
                            LayoutCachedLeft =2280
                            LayoutCachedTop =15
                            LayoutCachedWidth =3420
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6525
                    Width =525
                    Height =315
                    TabIndex =1
                    Name ="DBTitlest"
                    ControlSource ="DBTitlest"

                    LayoutCachedLeft =6525
                    LayoutCachedWidth =7050
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5295
                            Width =1080
                            Height =315
                            Name ="Label2"
                            Caption ="DB Title St:"
                            LayoutCachedLeft =5295
                            LayoutCachedWidth =6375
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =8925
                    Width =375
                    Height =255
                    TabIndex =2
                    Name ="Check4"
                    ControlSource ="DBTitleTrust"

                    LayoutCachedLeft =8925
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7410
                            Width =1395
                            Height =315
                            Name ="Label5"
                            Caption ="DB Title Trust"
                            LayoutCachedLeft =7410
                            LayoutCachedWidth =8805
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    Left =15
                    Top =435
                    Width =8790
                    Height =375
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="if this is blank then the VIN does not match any VIN in DB"
                    LayoutCachedLeft =15
                    LayoutCachedTop =435
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =810
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =2025
                    Height =315
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"

                    LayoutCachedWidth =2025
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
