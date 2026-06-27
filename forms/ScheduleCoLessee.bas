Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6900
    DatasheetFontHeight =11
    ItemSuffix =49
    Left =3420
    Top =4890
    Right =13425
    Bottom =11010
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x799764523202e540
    End
    RecordSource ="ScheduleCoLessee"
    Caption ="Co/Sub Lessee"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    SplitFormOrientation =3
    SplitFormOrientation =3
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
            Height =405
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =4455
                    Height =405
                    FontSize =14
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="Schedule Co/Sub Lessee Manager"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =405
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3780
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =5700
                    Top =120
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =5
                    ForeColor =5026082
                    Name ="ScheduleID"
                    ControlSource ="ScheduleID"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =120
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =825
                    Width =3945
                    Height =315
                    ColumnWidth =4335
                    FontSize =10
                    TabIndex =2
                    Name ="CoLesseeTaxID"
                    ControlSource ="CoLesseeTaxID"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =825
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1140
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =825
                            Width =1530
                            Height =315
                            FontSize =10
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label5"
                            Caption ="TaxID"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =825
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1140
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3660
                    Top =1200
                    Width =1905
                    Height =315
                    FontSize =9
                    TabIndex =4
                    Name ="SaveCoLessee"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save this Co Lessee"
                    GroupTable =1

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1515
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackColor =9732411
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =450
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =6
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="INDX"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =450
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =780
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =7
                    ForeColor =5026082
                    Name ="FAID"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =780
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =990
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5760
                    Left =1620
                    Top =75
                    Width =3945
                    Height =315
                    FontSize =10
                    BackColor =16249583
                    Name ="CoLesseeID"
                    ControlSource ="CoLesseeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientID, clientCompanyName AS CoLessee, clientTaxID As TaxID, ClientGrou"
                        "pId FROM  Clients WHERE ClientGroupID= 60 Order By clientCompanyName"
                    ColumnWidths ="0;3168;1152;0"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =75
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =30
                    Top =75
                    Width =1530
                    Height =315
                    FontSize =10
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label29"
                    Caption ="Co/Sub Lessee "
                    GroupTable =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =75
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =390
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =450
                    Width =3945
                    Height =315
                    FontSize =10
                    TabIndex =1
                    BackColor =16249583
                    Name ="CoLesseeLabel"
                    ControlSource ="CoLesseeLabel"
                    RowSourceType ="Value List"
                    RowSource ="\"Lessee\";\"Co Lessee\";\"Sub Lessee\";\"Obligator\";\"Co Obligator\""
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =450
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =450
                            Width =1530
                            Height =315
                            FontSize =10
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4"
                            Caption ="Leesee Type "
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =450
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =30
                    Top =1200
                    Width =1530
                    Height =315
                    Name ="EmptyCell37"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =30
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =135
                    Top =1680
                    Width =5445
                    Height =1725
                    FontSize =10
                    TabIndex =8
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="CoLesseeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ScheduleCoLessee.INDX, ScheduleCoLessee.CoLesseeLabel, Clients.clientComp"
                        "anyName, ScheduleCoLessee.ScheduleId, Clients.clientID FROM ScheduleCoLessee INN"
                        "ER JOIN Clients ON ScheduleCoLessee.CoLesseeID = Clients.clientID WHERE (((Sched"
                        "uleCoLessee.ScheduleId)=[Forms]![Schedule_Form![SchID]));"
                    ColumnWidths ="0;1440;3600;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =135
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =3405
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =1200
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =3
                    Name ="AddNew"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Add a Co Lessee"
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1515
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =16777215
                    HoverColor =42978
                    PressedColor =0
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3585
                    Top =3435
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =9
                    Name ="DeleteMe"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Delete the selected Co-Lessee"

                    LayoutCachedLeft =3585
                    LayoutCachedTop =3435
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =3750
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackColor =1643706
                    BorderColor =16777215
                    HoverColor =42978
                    PressedColor =1643706
                    QuickStyle =30
                    QuickStyleMask =-885
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
' See "ScheduleCoLessee.cls"
