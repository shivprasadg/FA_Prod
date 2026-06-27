Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6818
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =600
    Top =10485
    Right =9960
    Bottom =14610
    RecSrcDt = Begin
        0xc71ea06ed661e540
    End
    RecordSource ="SELECT ScheduleCoLessee.ScheduleID, ScheduleCoLessee.CoLesseeID, ScheduleCoLesse"
        "e.CoLesseeLabel, ScheduleCoLessee.CoLesseeTaxID, ScheduleCoLessee.HideOnCOA, Sch"
        "eduleCoLessee.INDX FROM ScheduleCoLessee;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
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
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4230
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2190
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="ScheduleID"
                    ControlSource ="ScheduleID"

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2190
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =2400
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =1185
                    Width =3945
                    Height =315
                    FontSize =10
                    TabIndex =3
                    BorderColor =12632256
                    Name ="CoLesseeTaxID"
                    ControlSource ="CoLesseeTaxID"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1185
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =1185
                            Width =1530
                            Height =315
                            FontSize =10
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label5"
                            Caption ="TaxID"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =1185
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1500
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =3660
                    Top =1560
                    Width =1905
                    Height =315
                    FontSize =9
                    TabIndex =5
                    Name ="SaveCoLessee"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save this Co Lessee"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1875
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =9732411
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =616409
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =1643706
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2520
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =8
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="INDX"

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =2730
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2850
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =9
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="FAID"

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2850
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =3060
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    Top =810
                    Width =3945
                    Height =315
                    FontSize =10
                    TabIndex =1
                    BackColor =16249583
                    BorderColor =14277081
                    Name ="CoLesseeID"
                    ControlSource ="CoLesseeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientID, clientCompanyName AS CoLessee, clientTaxID As TaxID, ClientGrou"
                        "pId FROM  Clients WHERE ClientGroupID= 27 ORDER BY clientCompanyName"
                    ColumnWidths ="0;3168;1152;0"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =810
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1125
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =30
                    Top =810
                    Width =1530
                    Height =315
                    FontSize =10
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label29"
                    Caption ="Co/Sub Lessee "
                    GroupTable =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =810
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1125
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =435
                    Width =3945
                    Height =315
                    FontSize =10
                    BackColor =16249583
                    BorderColor =14277081
                    Name ="CoLesseeLabel"
                    ControlSource ="CoLesseeLabel"
                    RowSourceType ="Value List"
                    RowSource ="\"Lessee\";\"Co-Lessee\";\"Sublessee\";\"Co-Obligor\""
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =1620
                    LayoutCachedTop =435
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =750
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =435
                            Width =1530
                            Height =315
                            FontSize =10
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4"
                            Caption ="Leesee Type "
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =435
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =750
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =30
                    Top =1560
                    Width =1530
                    Height =315
                    Name ="EmptyCell37"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =30
                    LayoutCachedTop =1560
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1875
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =45
                    Top =2010
                    Width =5445
                    Height =1725
                    FontSize =10
                    TabIndex =6
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="CoLesseeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Sch.INDX, Sch.CoLesseeLabel, Clients.clientCompanyName, Sch.ScheduleId, C"
                        "lients.clientID FROM ScheduleCoLessee AS Sch INNER JOIN Clients ON Sch.CoLesseeI"
                        "D = Clients.clientID WHERE (((Sch.ScheduleId)=1376)) ORDER BY Clients.clientComp"
                        "anyName;"
                    ColumnWidths ="0;1440;3600;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =45
                    LayoutCachedTop =2010
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =3735
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =1620
                    Top =1560
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =4
                    Name ="AddNew"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Add a Co Lessee"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1875
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =6
                    BackTint =100.0
                    BackShade =50.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =42978
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =3495
                    Top =3765
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =10
                    Name ="DeleteMe"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Delete the selected Co-Lessee"
                    GridlineColor =10921638

                    LayoutCachedLeft =3495
                    LayoutCachedTop =3765
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =4080
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =42978
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =1643706
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-885
                End
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
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5625
                    Top =810
                    Width =1155
                    Height =315
                    FontSize =10
                    TabIndex =2
                    BackColor =16249583
                    BorderColor =14277081
                    Name ="HideOnCOA"
                    ControlSource ="HideOnCOA"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Show\";-1;\"Hide\""
                    ColumnWidths ="0;7200"
                    FontName ="Segoe UI"
                    Format ="True/False"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =5625
                    LayoutCachedTop =810
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1125
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5625
                    Top =1185
                    Width =1155
                    Height =315
                    Name ="EmptyCell20"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5625
                    LayoutCachedTop =1185
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5625
                    Top =1560
                    Width =1155
                    Height =315
                    Name ="EmptyCell21"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5625
                    LayoutCachedTop =1560
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1875
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5625
                    Top =435
                    Width =1155
                    Height =315
                    FontSize =10
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label0"
                    Caption ="Hide On COA"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5625
                    LayoutCachedTop =435
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =750
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
    End
End
CodeBehindForm
' See "MasterLease_Guarentees.cls"
