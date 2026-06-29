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
    Width =12356
    DatasheetFontHeight =11
    ItemSuffix =19
    Left =9165
    Top =6420
    Right =25380
    Bottom =16515
    RecSrcDt = Begin
        0x03bf33f7d298e540
    End
    RecordSource ="Select * From SchGrp WHERE  sGrpId=1862"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin FormHeader
            Height =2130
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =85.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =1395
                    Top =30
                    Width =8520
                    Height =405
                    FontSize =16
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="FormLabel"
                    Caption ="Adding Group C to Schedule "
                    GridlineColor =0
                    LayoutCachedLeft =1395
                    LayoutCachedTop =30
                    LayoutCachedWidth =9915
                    LayoutCachedHeight =435
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =300
                    Top =120
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label169"
                    Caption ="Group No"
                    GridlineColor =0
                    LayoutCachedLeft =300
                    LayoutCachedTop =120
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =390
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =300
                    Top =420
                    Width =900
                    Height =540
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpNO"
                    ValidationText ="Group Number is required"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =38
                    GridlineColor =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =420
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =960
                    ColumnStart =1
                    ColumnEnd =1
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
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7650
                    Top =120
                    Width =1200
                    FontSize =10
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="SchID"
                    GridlineColor =0

                    LayoutCachedLeft =7650
                    LayoutCachedTop =120
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =7620
                    Top =960
                    Width =1575
                    Height =450
                    TabIndex =4
                    Name ="MakeGroup"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Insert this NEW Group"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7620
                    LayoutCachedTop =960
                    LayoutCachedWidth =9195
                    LayoutCachedHeight =1410
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =9732411
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =616409
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =1643706
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =1350
                    Top =420
                    Width =5970
                    Height =240
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label224"
                    Caption ="Reason for Adding Group"
                    GridlineColor =0
                    LayoutCachedLeft =1350
                    LayoutCachedTop =420
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =660
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =1515
                    Width =6855
                    Height =293
                    TabIndex =6
                    BackColor =13434879
                    ForeColor =0
                    Name ="SgrpDesc"
                    ControlSource ="SgrpDesc"
                    StatusBarText ="discription for the RAL Attach Letter"
                    GridlineColor =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =1515
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =1808
                    BackThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =85
                            Left =300
                            Top =1275
                            Width =6855
                            Height =218
                            FontSize =8
                            FontWeight =700
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label418"
                            Caption ="Group Equipment Description for RAL Attachment"
                            GridlineColor =0
                            LayoutCachedLeft =300
                            LayoutCachedTop =1275
                            LayoutCachedWidth =7155
                            LayoutCachedHeight =1493
                            ThemeFontIndex =-1
                            BackThemeColorIndex =2
                            BackTint =60.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1350
                    Top =660
                    Width =5835
                    Height =540
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="Reason"
                    StatusBarText ="blcd source accepted or in-service"
                    GridlineColor =0

                    LayoutCachedLeft =1350
                    LayoutCachedTop =660
                    LayoutCachedWidth =7185
                    LayoutCachedHeight =1200
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7368
                    Top =456
                    Width =3696
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="UnitGroupID"
                    ControlSource ="=\"SchID: \" & [SchID] & \"  GrpID: \" & [SGrpID] & \" BankID: \" & [AssigneeBan"
                        "kID]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7368
                    LayoutCachedTop =456
                    LayoutCachedWidth =11064
                    LayoutCachedHeight =696
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    PictureType =1
                    Left =9360
                    Top =960
                    Width =2880
                    Height =450
                    TabIndex =5
                    Name ="CloneSchNGroup"
                    Caption ="Clone This SCH and Group"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Insert this NEW Group"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9360
                    LayoutCachedTop =960
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =1410
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =7
                    BackTint =100.0
                    BackShade =75.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =7
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =7
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10428
                    Top =1500
                    Width =1860
                    Height =285
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =0
                    Name ="NewSchNo"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10428
                    LayoutCachedTop =1500
                    LayoutCachedWidth =12288
                    LayoutCachedHeight =1785
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =9360
                    Top =1500
                    Width =1050
                    Height =285
                    FontSize =10
                    BorderColor =0
                    Name ="Label4"
                    Caption ="New Sch No"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9360
                    LayoutCachedTop =1500
                    LayoutCachedWidth =10410
                    LayoutCachedHeight =1785
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10428
                    Top =1800
                    Width =1860
                    Height =288
                    TabIndex =8
                    BorderColor =12632256
                    ForeColor =0
                    Name ="NewGrpNo"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10428
                    LayoutCachedTop =1800
                    LayoutCachedWidth =12288
                    LayoutCachedHeight =2088
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =9360
                    Top =1800
                    Width =1050
                    Height =288
                    FontSize =10
                    BorderColor =0
                    Name ="Label6"
                    Caption ="New Grp No"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9360
                    LayoutCachedTop =1800
                    LayoutCachedWidth =10410
                    LayoutCachedHeight =2088
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =120
                    Width =1200
                    FontSize =10
                    TabIndex =9
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="SourceGrpID"
                    GridlineColor =0

                    LayoutCachedLeft =8940
                    LayoutCachedTop =120
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin Section
            Height =10476
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1815
                    Top =330
                    Width =1380
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label1"
                    Caption ="Lease Type"
                    GridlineColor =0
                    LayoutCachedLeft =1815
                    LayoutCachedTop =330
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =3300
                    Top =870
                    Width =750
                    Height =210
                    FontSize =9
                    BorderColor =0
                    Name ="Label4257"
                    Caption ="Group"
                    GridlineColor =0
                    LayoutCachedLeft =3300
                    LayoutCachedTop =870
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =1080
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =6855
                    Top =870
                    Width =1319
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label4253"
                    Caption ="for Group"
                    GridlineColor =0
                    LayoutCachedLeft =6855
                    LayoutCachedTop =870
                    LayoutCachedWidth =8174
                    LayoutCachedHeight =1095
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =75
                    Top =330
                    Width =1620
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="grpfaidlbl"
                    Caption ="FAID"
                    GridlineColor =0
                    LayoutCachedLeft =75
                    LayoutCachedTop =330
                    LayoutCachedWidth =1695
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3300
                    Top =540
                    Width =750
                    Height =330
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="UnitsInGroup"
                    ValidationRule ="Is Not Null And >0"
                    ValidationText ="No of Units is a required field on the group record."
                    GridlineColor =0

                    LayoutCachedLeft =3300
                    LayoutCachedTop =540
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =870
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
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5250
                    Top =540
                    Width =1499
                    Height =330
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Total Cost to be syndicated"
                    ValidationRule =">=0"
                    ValidationText ="Equipment cost  must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Total Cost to be syndicated"
                    GridlineColor =0

                    LayoutCachedLeft =5250
                    LayoutCachedTop =540
                    LayoutCachedWidth =6749
                    LayoutCachedHeight =870
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6181
                    Top =4035
                    Width =510
                    FontSize =10
                    TabIndex =2
                    BackColor =-2147483600
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="ExchangeITMonth"
                    ControlSource ="ExchangeITMonth"
                    StatusBarText ="Exchange Month Not notice date"
                    ValidationRule =">=0"
                    ValidationText ="Must be >= 0"
                    DefaultValue ="0"
                    ControlTipText ="Exchange Month Not notice date"
                    BottomPadding =150
                    GridlineColor =0

                    LayoutCachedLeft =6181
                    LayoutCachedTop =4035
                    LayoutCachedWidth =6691
                    LayoutCachedHeight =4275
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5190
                            Top =4035
                            Width =930
                            Height =240
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="Label14"
                            Caption ="@ Months"
                            BottomPadding =150
                            GridlineColor =0
                            LayoutCachedLeft =5190
                            LayoutCachedTop =4035
                            LayoutCachedWidth =6120
                            LayoutCachedHeight =4275
                            ColumnStart =3
                            ColumnEnd =3
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =90
                    Top =2025
                    Width =960
                    Height =360
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeaseTermPrimary"
                    ControlSource ="LeaseTermPrimary"
                    StatusBarText ="primary term in months"
                    GridlineColor =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =2025
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =2385
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =90
                            Top =2025
                            Width =960
                            Height =240
                            FontSize =10
                            ForeColor =6108695
                            Name ="COmmLBL"
                            Caption =" Term"
                            GridlineColor =0
                            LayoutCachedLeft =90
                            LayoutCachedTop =2025
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =2265
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =50.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8293
                    Top =540
                    Width =1320
                    Height =330
                    TabIndex =4
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LRFGroup"
                    ControlSource ="LRFGroup"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =8293
                    LayoutCachedTop =540
                    LayoutCachedWidth =9613
                    LayoutCachedHeight =870
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6854
                    Top =540
                    Width =1319
                    Height =330
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeasePymtForGroup"
                    ControlSource ="LeasePymtForGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Rent Amounts must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6854
                    LayoutCachedTop =540
                    LayoutCachedWidth =8173
                    LayoutCachedHeight =870
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1815
                    Top =540
                    Width =1365
                    Height =330
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeaseType"
                    ControlSource ="LeaseType"
                    RowSourceType ="Value List"
                    RowSource ="\"FMV\";T.R.A.C.;\"Split T.R.A.C.\";\"CAPITAL\";REMARKET;REPORTING"
                    Tag ="Protected"
                    GridlineColor =0

                    LayoutCachedLeft =1815
                    LayoutCachedTop =540
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =870
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1815
                            Top =540
                            Width =1365
                            Height =210
                            FontSize =10
                            BorderColor =0
                            ForeColor =6108695
                            Name ="ltypelbl"
                            Caption ="Lease Type"
                            GridlineColor =0
                            LayoutCachedLeft =1815
                            LayoutCachedTop =540
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =750
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1995
                    Top =4035
                    Width =3166
                    FontSize =10
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ExchangeITType"
                    ControlSource ="ExchangeITType"
                    RowSourceType ="Value List"
                    RowSource ="Specified Mo;Any Time-90 day Notice;\"Quarterly starting at...\""
                    BottomPadding =150
                    GridlineColor =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1995
                    LayoutCachedTop =4035
                    LayoutCachedWidth =5161
                    LayoutCachedHeight =4275
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1140
                            Top =4035
                            Width =750
                            Height =240
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="EXTypLBL"
                            Caption ="Type of "
                            BottomPadding =150
                            GridlineColor =0
                            LayoutCachedLeft =1140
                            LayoutCachedTop =4035
                            LayoutCachedWidth =1890
                            LayoutCachedHeight =4275
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1185
                    Top =2025
                    Width =750
                    Height =360
                    TabIndex =8
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeaseTermNotice"
                    ControlSource ="LeaseTermNotice"
                    ControlTipText ="Prior notice applies to ExchangeIt and all extension options"
                    GridlineColor =0

                    LayoutCachedLeft =1185
                    LayoutCachedTop =2025
                    LayoutCachedWidth =1935
                    LayoutCachedHeight =2385
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
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =2
                    Left =3509
                    Top =1995
                    Width =1079
                    Height =360
                    TabIndex =9
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeasePaidAdvArr"
                    ControlSource ="LeasePaidAdvArr"
                    RowSourceType ="Value List"
                    RowSource ="\"Adv\";\"Arr\""
                    ColumnWidths ="720"
                    Tag ="Protected"

                    LayoutCachedLeft =3509
                    LayoutCachedTop =1995
                    LayoutCachedWidth =4588
                    LayoutCachedHeight =2355
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2070
                    Top =1995
                    Width =1319
                    Height =360
                    TabIndex =10
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeasePymtFreq"
                    ControlSource ="LeasePymtFreq"
                    RowSourceType ="Value List"
                    RowSource ="M;MONTHLY;Q;QUARTERLY;S;SEMI-ANN;A;ANNUAL"
                    ColumnWidths ="0;720"
                    StatusBarText ="pmt period"
                    Tag ="Protected"

                    LayoutCachedLeft =2070
                    LayoutCachedTop =1995
                    LayoutCachedWidth =3389
                    LayoutCachedHeight =2355
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =4035
                    Width =960
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ExchangeITGroup"
                    ControlSource ="ExchangeITGroup"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;576"
                    Tag ="Protected"
                    Format ="True/False"
                    BottomPadding =150

                    LayoutCachedLeft =75
                    LayoutCachedTop =4035
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =4275
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =3690
                    Width =12356
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483600
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label4042"
                    Caption ="Exchange IT"
                    GridlineColor =0
                    LayoutCachedTop =3690
                    LayoutCachedWidth =12356
                    LayoutCachedHeight =3930
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    Left =90
                    Top =2385
                    Width =960
                    Height =210
                    FontSize =9
                    BorderColor =0
                    Name ="Label4114"
                    Caption ="Committed"
                    GridlineColor =0
                    LayoutCachedLeft =90
                    LayoutCachedTop =2385
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =2595
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =2160
                    Left =5190
                    Top =1785
                    Width =4980
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SgrpCommInd"
                    ControlSource ="SgrpCommInd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GridlineColor =0

                    LayoutCachedLeft =5190
                    LayoutCachedTop =1785
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =2115
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1185
                    Top =1785
                    Width =765
                    Height =240
                    FontSize =10
                    ForeColor =6108695
                    Name ="Label4116"
                    Caption ="Notice"
                    GridlineColor =0
                    LayoutCachedLeft =1185
                    LayoutCachedTop =1785
                    LayoutCachedWidth =1950
                    LayoutCachedHeight =2025
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1185
                    Top =2400
                    Width =750
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label4117"
                    Caption ="Months"
                    GridlineColor =0
                    LayoutCachedLeft =1185
                    LayoutCachedTop =2400
                    LayoutCachedWidth =1935
                    LayoutCachedHeight =2625
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =3300
                    Top =330
                    Width =750
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Nofounitslbl"
                    Caption ="Units In"
                    GridlineColor =0
                    LayoutCachedLeft =3300
                    LayoutCachedTop =330
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2070
                    Top =1785
                    Width =1319
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="PisLBL"
                    Caption ="Payments are"
                    LayoutCachedLeft =2070
                    LayoutCachedTop =1785
                    LayoutCachedWidth =3389
                    LayoutCachedHeight =1995
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =3509
                    Top =1785
                    Width =1079
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="MdeinLBL"
                    Caption ="Paid in"
                    LayoutCachedLeft =3509
                    LayoutCachedTop =1785
                    LayoutCachedWidth =4588
                    LayoutCachedHeight =1995
                    RowStart =1
                    RowEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =5235
                    Top =330
                    Width =1499
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="grpoeclbl"
                    Caption ="Group OEC"
                    GridlineColor =0
                    LayoutCachedLeft =5235
                    LayoutCachedTop =330
                    LayoutCachedWidth =6734
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =6854
                    Top =330
                    Width =1319
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="LPMTLBL"
                    Caption ="Lease Payment"
                    LayoutCachedLeft =6854
                    LayoutCachedTop =330
                    LayoutCachedWidth =8173
                    LayoutCachedHeight =540
                    RowStart =3
                    RowEnd =3
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =8293
                    Top =330
                    Width =1320
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="LRFLBL"
                    Caption ="LRF"
                    LayoutCachedLeft =8293
                    LayoutCachedTop =330
                    LayoutCachedWidth =9613
                    LayoutCachedHeight =540
                    RowStart =3
                    RowEnd =3
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =2
                    ListWidth =720
                    Left =4170
                    Top =540
                    Width =900
                    Height =330
                    TabIndex =13
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGRPPLB"
                    ControlSource ="SGRPPLB"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    ColumnWidths ="720"
                    DefaultValue ="No"
                    Tag ="Protected"
                    ControlTipText ="Is this a Purchase Lease Back (PLB)?"
                    Format ="Yes/No"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4170
                    LayoutCachedTop =540
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =870
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =4170
                    Top =330
                    Width =900
                    Height =210
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label4261"
                    Caption ="PLB"
                    GridlineColor =0
                    LayoutCachedLeft =4170
                    LayoutCachedTop =330
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =5205
                    Top =1485
                    Width =4980
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =8872517
                    Name ="CommBaselbl"
                    Caption ="BLCD Basis"
                    GridlineColor =0
                    LayoutCachedLeft =5205
                    LayoutCachedTop =1485
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =1725
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =135
                    Top =3285
                    Width =960
                    Height =315
                    FontSize =10
                    TabIndex =14
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeaseExtOption1Term"
                    ControlSource ="LeaseExtOption1Term"
                    StatusBarText ="1st option in months"
                    ControlTipText ="1st option in months"
                    GridlineColor =0

                    LayoutCachedLeft =135
                    LayoutCachedTop =3285
                    LayoutCachedWidth =1095
                    LayoutCachedHeight =3600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
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
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1230
                    Top =3285
                    Width =1005
                    Height =315
                    FontSize =10
                    TabIndex =15
                    BorderColor =12632256
                    ForeColor =0
                    Name ="LeaseExtOption1Rent"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="1st option lrf"
                    ValidationText ="Rent must be >=0"
                    DefaultValue ="0"
                    Tag ="UAT"

                    LayoutCachedLeft =1230
                    LayoutCachedTop =3285
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =3600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2370
                    Top =3285
                    Width =885
                    Height =315
                    FontSize =10
                    TabIndex =16
                    BorderColor =12632256
                    ForeColor =0
                    Name ="MTMRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MTM rent"
                    ValidationText ="Rent must be >=0"
                    DefaultValue ="0"
                    Tag ="UAT"
                    ControlTipText ="Rent to be used as Month To Month Rent"

                    LayoutCachedLeft =2370
                    LayoutCachedTop =3285
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =120
                    Top =3045
                    Width =975
                    Height =240
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label2514"
                    Caption ="Ext'd Term"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =3045
                    LayoutCachedWidth =1095
                    LayoutCachedHeight =3285
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1215
                    Top =3045
                    Width =1035
                    Height =240
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label2544"
                    Caption ="Ext'd Rent"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1215
                    LayoutCachedTop =3045
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =3285
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =2730
                    Width =4980
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3990"
                    Caption ="Extend IT"
                    GridlineColor =0
                    LayoutCachedTop =2730
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =2970
                    ThemeFontIndex =-1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =5205
                    Top =3045
                    Width =4980
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PerDiemStartsOn"
                    ControlSource ="PerDiemStartsOn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GridlineColor =0

                    LayoutCachedLeft =5205
                    LayoutCachedTop =3045
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =3375
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Width =12356
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =8872517
                    Name ="Label182"
                    Caption ="Structure"
                    GridlineColor =0
                    LayoutCachedWidth =12356
                    LayoutCachedHeight =240
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =1485
                    Width =4980
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =8872517
                    Name ="Label183"
                    Caption ="Terms"
                    GridlineColor =0
                    LayoutCachedTop =1485
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =1725
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =5205
                    Top =2730
                    Width =4980
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =10855845
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label185"
                    Caption ="Per Diem Starts On"
                    GridlineColor =0
                    LayoutCachedLeft =5205
                    LayoutCachedTop =2730
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =2970
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2370
                    Top =3045
                    Width =885
                    Height =240
                    FontSize =10
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label213"
                    Caption ="MTM Rent"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2370
                    LayoutCachedTop =3045
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3285
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =75
                    Top =540
                    Width =1620
                    Height =330
                    TabIndex =18
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpFAID"
                    ControlSource ="SGrpFAID"
                    StatusBarText ="Enter ONLY if different from schedule FAID"
                    Tag ="Protected"
                    ControlTipText ="Entere ONLY if different from schedule FAID"
                    GridlineColor =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =540
                    LayoutCachedWidth =1695
                    LayoutCachedHeight =870
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
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1170
                    Top =5160
                    Width =1140
                    Height =255
                    FontSize =9
                    TabIndex =19
                    BorderColor =12632256
                    Name ="SGrpSalesIndex"
                    ControlSource ="SGrpSalesIndex"
                    RowSourceType ="Value List"
                    RowSource ="\"SWAP\";\"T-BILL\""
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =5160
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =5415
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =210
                            Top =5160
                            Width =900
                            Height =255
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label281"
                            Caption ="Index"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =210
                            LayoutCachedTop =5160
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =5415
                            RowStart =7
                            RowEnd =7
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1170
                    Top =5475
                    Width =1140
                    FontSize =9
                    TabIndex =20
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="SGrpSalesIndDt"
                    ControlSource ="SGrpSalesIndDt"
                    Format ="Short Date"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =5475
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =5715
                    RowStart =8
                    RowEnd =8
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =210
                            Top =5475
                            Width =900
                            Height =240
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label282"
                            Caption ="Date"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =210
                            LayoutCachedTop =5475
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =5715
                            RowStart =8
                            RowEnd =8
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1170
                    Top =5775
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =21
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="SGprSalesIndYr"
                    ControlSource ="SGprSalesIndYr"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =5775
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =6060
                    RowStart =9
                    RowEnd =9
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =210
                            Top =5775
                            Width =900
                            Height =285
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label283"
                            Caption ="Term/Year"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =210
                            LayoutCachedTop =5775
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =6060
                            RowStart =9
                            RowEnd =9
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1170
                    Top =6105
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =22
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="SgrpSalesIndRt"
                    ControlSource ="SgrpSalesIndRt"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =6105
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =6390
                    RowStart =10
                    RowEnd =10
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =210
                            Top =6105
                            Width =900
                            Height =285
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label284"
                            Caption ="Rate"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =210
                            LayoutCachedTop =6105
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =6390
                            RowStart =10
                            RowEnd =10
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =5160
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =23
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="SalesSpread"
                    ControlSource ="=[SgrpSalesYield]-[SgrpSalesIndRt]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =5160
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =5445
                    RowStart =11
                    RowEnd =11
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2460
                            Top =5160
                            Width =900
                            Height =285
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label286"
                            Caption ="Spread"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2460
                            LayoutCachedTop =5160
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =5445
                            RowStart =11
                            RowEnd =11
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =5505
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =24
                    BorderColor =12632256
                    ForeColor =5855577
                    Name ="SgrpSalesYield"
                    ControlSource ="SgrpSalesYield"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =5505
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =5790
                    RowStart =12
                    RowEnd =12
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2460
                            Top =5505
                            Width =900
                            Height =285
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label285"
                            Caption ="Yield"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2460
                            LayoutCachedTop =5505
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =5790
                            RowStart =12
                            RowEnd =12
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =5850
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =25
                    BorderColor =12632256
                    Name ="Bonus"
                    ControlSource ="SGrpSalesBonus"
                    RowSourceType ="Value List"
                    RowSource ="\"NONE\";\"30%\";\"40%\";\"50%\";\"100%\""
                    StatusBarText ="Bonus Depreciation"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =5850
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =6135
                    RowStart =13
                    RowEnd =13
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    ForeShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =210
                    Top =4785
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label1505"
                    Caption ="LO/Target"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedTop =4785
                    LayoutCachedWidth =1350
                    LayoutCachedHeight =5085
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2460
                    Top =5850
                    Width =900
                    Height =285
                    FontSize =9
                    BorderColor =0
                    ForeColor =0
                    Name ="Label299"
                    Caption ="Bonus"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =5850
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =6135
                    RowStart =13
                    RowEnd =13
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =4470
                    Width =12356
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3634"
                    Caption ="Economics Detail"
                    GridlineStyleBottom =1
                    GridlineColor =0
                    LayoutCachedTop =4470
                    LayoutCachedWidth =12356
                    LayoutCachedHeight =4710
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =6540
                    Width =12356
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =4754549
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3683"
                    Caption ="Mileage Allowances"
                    GridlineStyleBottom =1
                    GridlineColor =0
                    LayoutCachedTop =6540
                    LayoutCachedWidth =12356
                    LayoutCachedHeight =6780
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =7020
                    Width =2832
                    Height =270
                    FontSize =10
                    BackColor =14277081
                    BorderColor =0
                    ForeColor =0
                    Name ="Label95"
                    Caption ="Allowance Method"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =7020
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =7290
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3012
                    Top =7020
                    Width =240
                    Height =270
                    Name ="EmptyCell3745"
                    GroupTable =1
                    LayoutCachedLeft =3012
                    LayoutCachedTop =7020
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =7290
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3312
                    Top =7020
                    Width =372
                    Height =270
                    Name ="EmptyCell3746"
                    GroupTable =1
                    LayoutCachedLeft =3312
                    LayoutCachedTop =7020
                    LayoutCachedWidth =3684
                    LayoutCachedHeight =7290
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3756
                    Top =7020
                    Width =1320
                    Height =270
                    FontSize =9
                    BackColor =14277081
                    BorderColor =0
                    ForeColor =0
                    Name ="Label390"
                    Caption ="Charged After"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =3756
                    LayoutCachedTop =7020
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =7290
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5148
                    Top =7020
                    Width =1188
                    Height =270
                    FontSize =9
                    BackColor =14277081
                    BorderColor =0
                    ForeColor =0
                    Name ="Label391"
                    Caption ="Charged Amt"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =5148
                    LayoutCachedTop =7020
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =7290
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =6408
                    Top =7020
                    Width =1812
                    Height =270
                    FontSize =9
                    BackColor =14277081
                    BorderColor =0
                    ForeColor =0
                    Name ="Label360"
                    Caption ="Equipment Type"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =6408
                    LayoutCachedTop =7020
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =7290
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1632
                    Top =7368
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =26
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGPRU/A"
                    ControlSource ="SGPRU/A"
                    RowSourceType ="Value List"
                    RowSource ="PER UNIT;IN AGG"
                    FontName ="Segoe UI"
                    EventProcPrefix ="SGPRU_A"
                    GroupTable =1

                    LayoutCachedLeft =1632
                    LayoutCachedTop =7368
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =7644
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =7368
                            Width =1440
                            Height =276
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label46"
                            Caption ="Overcharge is"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =7368
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =7644
                            RowStart =1
                            RowEnd =1
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
                Begin EmptyCell
                    Left =3012
                    Top =7368
                    Width =240
                    Height =276
                    Name ="EmptyCell3700"
                    GroupTable =1
                    LayoutCachedLeft =3012
                    LayoutCachedTop =7368
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =7644
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =7368
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =27
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow1"
                    ControlSource ="SGprAllow1"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =1

                    LayoutCachedLeft =3756
                    LayoutCachedTop =7368
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =7644
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =7368
                            Width =372
                            Height =276
                            FontSize =9
                            BackColor =6108695
                            BorderColor =0
                            Name ="Label152"
                            Caption ="1"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =3312
                            LayoutCachedTop =7368
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =7644
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =7368
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =28
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov1"
                    ControlSource ="SGrpov1"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =5148
                    LayoutCachedTop =7368
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =7644
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6408
                    Top =7368
                    Width =1812
                    Height =276
                    FontSize =10
                    TabIndex =29
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SgrpEq1"
                    ControlSource ="SgrpEq1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                        "lookup_groupTypes.groupType;"
                    StatusBarText ="For use when group has different equipment to specify over charge for a type"
                    FontName ="Segoe UI"
                    ControlTipText ="For use when group has different equipment to specify over charge for a type"
                    GroupTable =1

                    LayoutCachedLeft =6408
                    LayoutCachedTop =7368
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =7644
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =2
                    Left =1632
                    Top =7716
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =30
                    BorderColor =12632256
                    ForeColor =0
                    Name ="AllowedMilesPeriod"
                    ControlSource ="AllowedMilesPeriod"
                    RowSourceType ="Value List"
                    RowSource ="P/YEAR;P/TERM"
                    FontName ="Segoe UI"
                    GroupTable =1
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =1632
                    LayoutCachedTop =7716
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =7992
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =7716
                            Width =1440
                            Height =276
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label125"
                            Caption ="Overage Period"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =7716
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =7992
                            RowStart =2
                            RowEnd =2
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
                Begin EmptyCell
                    Left =3012
                    Top =7716
                    Width =240
                    Height =276
                    Name ="EmptyCell3695"
                    GroupTable =1
                    LayoutCachedLeft =3012
                    LayoutCachedTop =7716
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =7992
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =7716
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =31
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow2"
                    ControlSource ="SGprAllow2"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =1

                    LayoutCachedLeft =3756
                    LayoutCachedTop =7716
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =7992
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =7716
                            Width =372
                            Height =276
                            FontSize =9
                            BackColor =6108695
                            BorderColor =0
                            Name ="Label153"
                            Caption ="2"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =3312
                            LayoutCachedTop =7716
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =7992
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =7716
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =32
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov2"
                    ControlSource ="SGrpov2"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =5148
                    LayoutCachedTop =7716
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =7992
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6408
                    Top =7716
                    Width =1812
                    Height =276
                    FontSize =10
                    TabIndex =33
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SgrpEq2"
                    ControlSource ="SgrpEq2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                        "lookup_groupTypes.groupType;"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6408
                    LayoutCachedTop =7716
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =7992
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =120
                    Top =8064
                    Height =276
                    Name ="EmptyCell3703"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =8064
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1632
                    Top =8064
                    Width =1320
                    Height =276
                    Name ="EmptyCell3704"
                    GroupTable =1
                    LayoutCachedLeft =1632
                    LayoutCachedTop =8064
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3012
                    Top =8064
                    Width =240
                    Height =276
                    Name ="EmptyCell3705"
                    GroupTable =1
                    LayoutCachedLeft =3012
                    LayoutCachedTop =8064
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =8064
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =34
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow3"
                    ControlSource ="SGprAllow3"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =1

                    LayoutCachedLeft =3756
                    LayoutCachedTop =8064
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =8064
                            Width =372
                            Height =276
                            FontSize =9
                            BackColor =6108695
                            BorderColor =0
                            Name ="Label154"
                            Caption ="3"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =3312
                            LayoutCachedTop =8064
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =8340
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =8064
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =35
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov3"
                    ControlSource ="SGrpov3"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =5148
                    LayoutCachedTop =8064
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6408
                    Top =8064
                    Width =1812
                    Height =276
                    FontSize =10
                    TabIndex =36
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SgrpEq3"
                    ControlSource ="SgrpEq3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                        "lookup_groupTypes.groupType;"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6408
                    LayoutCachedTop =8064
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =8340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9930
                    Top =7680
                    Width =2280
                    Height =600
                    FontSize =9
                    TabIndex =37
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="AllowedMilesNote"
                    ControlSource ="AllowedMilesNote"
                    StatusBarText ="special mi provision"
                    FontName ="Segoe UI"
                    GridlineColor =0

                    LayoutCachedLeft =9930
                    LayoutCachedTop =7680
                    LayoutCachedWidth =12210
                    LayoutCachedHeight =8280
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8310
                            Top =7680
                            Width =1560
                            Height =600
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="Label385"
                            Caption ="Other Mileage Provision"
                            GridlineColor =0
                            LayoutCachedLeft =8310
                            LayoutCachedTop =7680
                            LayoutCachedWidth =9870
                            LayoutCachedHeight =8280
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    Left =705
                    Top =8580
                    Width =11340
                    Name ="Line379"
                    GridlineColor =0
                    LayoutCachedLeft =705
                    LayoutCachedTop =8580
                    LayoutCachedWidth =12045
                    LayoutCachedHeight =8580
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    Top =8640
                    Width =12356
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3770"
                    Caption ="Hours Allowances"
                    GridlineStyleBottom =1
                    GridlineColor =0
                    LayoutCachedTop =8640
                    LayoutCachedWidth =12356
                    LayoutCachedHeight =8880
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =3
                    BackShade =25.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =9120
                    Width =2832
                    Height =270
                    FontSize =10
                    BorderColor =0
                    ForeColor =0
                    Name ="Label3813"
                    Caption ="Allowance Method"
                    GroupTable =2
                    LayoutCachedLeft =120
                    LayoutCachedTop =9120
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =9390
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =3012
                    Top =9120
                    Width =240
                    Height =270
                    Name ="EmptyCell3817"
                    GroupTable =2
                    LayoutCachedLeft =3012
                    LayoutCachedTop =9120
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =9390
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =3312
                    Top =9120
                    Width =372
                    Height =270
                    Name ="EmptyCell3818"
                    GroupTable =2
                    LayoutCachedLeft =3312
                    LayoutCachedTop =9120
                    LayoutCachedWidth =3684
                    LayoutCachedHeight =9390
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3756
                    Top =9120
                    Width =1320
                    Height =270
                    FontSize =9
                    BorderColor =0
                    ForeColor =0
                    Name ="Label3821"
                    Caption ="Charged After"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =3756
                    LayoutCachedTop =9120
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =9390
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5148
                    Top =9120
                    Width =1188
                    Height =270
                    FontSize =9
                    BorderColor =0
                    ForeColor =0
                    Name ="Label3823"
                    Caption ="Charged Amt"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =5148
                    LayoutCachedTop =9120
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =9390
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1632
                    Top =9468
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =38
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGPRHU/A"
                    ControlSource ="SGPRHU/A"
                    RowSourceType ="Value List"
                    RowSource ="PER UNIT;IN AGG"
                    FontName ="Segoe UI"
                    EventProcPrefix ="SGPRHU_A"
                    GroupTable =2

                    LayoutCachedLeft =1632
                    LayoutCachedTop =9468
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =9744
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =9468
                            Width =1440
                            Height =276
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label382"
                            Caption ="Overcharge is"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =120
                            LayoutCachedTop =9468
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =9744
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =3012
                    Top =9468
                    Width =240
                    Height =276
                    Name ="EmptyCell3786"
                    GroupTable =2
                    LayoutCachedLeft =3012
                    LayoutCachedTop =9468
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =9744
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =9468
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =39
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow4"
                    ControlSource ="SGprAllow4"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =2

                    LayoutCachedLeft =3756
                    LayoutCachedTop =9468
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =9744
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =9468
                            Width =372
                            Height =276
                            FontSize =9
                            BorderColor =0
                            Name ="Label3780"
                            Caption ="1"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =3312
                            LayoutCachedTop =9468
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =9744
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =6
                            BackShade =50.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =9468
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =40
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov4"
                    ControlSource ="SGrpov4"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =5148
                    LayoutCachedTop =9468
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =9744
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =2
                    Left =1632
                    Top =9816
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =41
                    BorderColor =12632256
                    ForeColor =0
                    Name ="AllowedHoursPeriod"
                    ControlSource ="AllowedHoursPeriod"
                    RowSourceType ="Value List"
                    RowSource ="P/YEAR;P/TERM"
                    FontName ="Segoe UI"
                    GroupTable =2
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =1632
                    LayoutCachedTop =9816
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =10092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =9816
                            Width =1440
                            Height =276
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label384"
                            Caption ="Overage Period"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =120
                            LayoutCachedTop =9816
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =10092
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =3012
                    Top =9816
                    Width =240
                    Height =276
                    Name ="EmptyCell3782"
                    GroupTable =2
                    LayoutCachedLeft =3012
                    LayoutCachedTop =9816
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =10092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =9816
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =42
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow5"
                    ControlSource ="SGprAllow5"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =2

                    LayoutCachedLeft =3756
                    LayoutCachedTop =9816
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =10092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =9816
                            Width =372
                            Height =276
                            FontSize =9
                            BorderColor =0
                            Name ="Label3809"
                            Caption ="2"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =3312
                            LayoutCachedTop =9816
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =10092
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =6
                            BackShade =50.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =9816
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =43
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov5"
                    ControlSource ="SGrpov5"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =5148
                    LayoutCachedTop =9816
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =10092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =10164
                    Height =276
                    Name ="EmptyCell3801"
                    GroupTable =2
                    LayoutCachedLeft =120
                    LayoutCachedTop =10164
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =10440
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =1632
                    Top =10164
                    Width =1320
                    Height =276
                    Name ="EmptyCell3802"
                    GroupTable =2
                    LayoutCachedLeft =1632
                    LayoutCachedTop =10164
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =10440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =3012
                    Top =10164
                    Width =240
                    Height =276
                    Name ="EmptyCell3803"
                    GroupTable =2
                    LayoutCachedLeft =3012
                    LayoutCachedTop =10164
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =10440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3756
                    Top =10164
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =44
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGprAllow6"
                    ControlSource ="SGprAllow6"
                    Format ="Standard"
                    StatusBarText ="Enter all mile overages before hrs"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    ControlTipText ="Enter all mile overages before hrs"
                    GroupTable =2

                    LayoutCachedLeft =3756
                    LayoutCachedTop =10164
                    LayoutCachedWidth =5076
                    LayoutCachedHeight =10440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =3312
                            Top =10164
                            Width =372
                            Height =276
                            FontSize =9
                            BorderColor =0
                            Name ="Label3810"
                            Caption ="3"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =3312
                            LayoutCachedTop =10164
                            LayoutCachedWidth =3684
                            LayoutCachedHeight =10440
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =6
                            BackShade =50.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5148
                    Top =10164
                    Width =1188
                    Height =276
                    FontSize =10
                    TabIndex =45
                    BorderColor =12632256
                    ForeColor =0
                    Name ="SGrpov6"
                    ControlSource ="SGrpov6"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =5148
                    LayoutCachedTop =10164
                    LayoutCachedWidth =6336
                    LayoutCachedHeight =10440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9930
                    Top =9780
                    Width =2280
                    Height =600
                    FontSize =9
                    TabIndex =46
                    BackColor =13434879
                    BorderColor =12632256
                    ForeColor =0
                    Name ="AllowedHoursNote"
                    ControlSource ="AllowedHoursNote"
                    StatusBarText ="special mi provision"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =9930
                    LayoutCachedTop =9780
                    LayoutCachedWidth =12210
                    LayoutCachedHeight =10380
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8505
                            Top =9780
                            Width =1365
                            Height =465
                            FontSize =9
                            BorderColor =0
                            ForeColor =0
                            Name ="Label388"
                            Caption ="Other Hours Provision"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =8505
                            LayoutCachedTop =9780
                            LayoutCachedWidth =9870
                            LayoutCachedHeight =10245
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Top =1785
                    Width =1064
                    Height =210
                    FontSize =10
                    ForeColor =6108695
                    Name ="Label0"
                    Caption =" Term Months"
                    GridlineColor =0
                    LayoutCachedTop =1785
                    LayoutCachedWidth =1064
                    LayoutCachedHeight =1995
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "GroupAdd_Form.cls"
