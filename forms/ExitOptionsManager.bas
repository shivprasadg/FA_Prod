Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =26438
    DatasheetFontHeight =11
    ItemSuffix =747
    Left =-32251
    Top =8430
    Right =-5551
    Bottom =22530
    ShortcutMenuBar ="SortClip"
    OrderBy ="[ExitOptionsManager].[Client]"
    RecSrcDt = Begin
        0xb5f83b8eb4fce540
    End
    RecordSource ="Select Distinct * From ExitOptionsGroup_Local Where AssetID > 0  AND DaysToEOL <"
        " 181 Order By DaysToEOL "
    Caption ="EXIT Options Dashboard"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2614
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =2419
                    Width =26424
                    Height =195
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =2419
                    LayoutCachedWidth =26424
                    LayoutCachedHeight =2614
                    TabIndex =30
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =20895
                    Top =2280
                    Width =780
                    Height =270
                    FontSize =9
                    Name ="New Unit_Label"
                    Caption ="AssetID"
                    EventProcPrefix ="New_Unit_Label"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20895
                    LayoutCachedTop =2280
                    LayoutCachedWidth =21675
                    LayoutCachedHeight =2550
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4875
                    Top =2280
                    Width =780
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4875
                    LayoutCachedTop =2280
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =2550
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5715
                    Top =2280
                    Width =885
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitUnitNum_Label"
                    Caption ="Unit #"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5715
                    LayoutCachedTop =2280
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =2550
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6660
                    Top =2280
                    Width =540
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitStatus_Label"
                    Caption ="Status"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6660
                    LayoutCachedTop =2280
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =2550
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16395
                    Top =2280
                    Width =1710
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitVIN_Label"
                    Caption ="Options Selected"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16395
                    LayoutCachedTop =2280
                    LayoutCachedWidth =18105
                    LayoutCachedHeight =2550
                    ColumnStart =18
                    ColumnEnd =20
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =13605
                    Top =2280
                    Width =645
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="UnitModel_Label"
                    Caption ="Until"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13605
                    LayoutCachedTop =2280
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =2550
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    Left =18165
                    Top =2280
                    Width =2664
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitCOASent_Label"
                    Caption ="New Schedule |Group"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18165
                    LayoutCachedTop =2280
                    LayoutCachedWidth =20829
                    LayoutCachedHeight =2550
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =23520
                    Top =2280
                    Width =870
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label122"
                    Caption ="Accepeted"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23520
                    LayoutCachedTop =2280
                    LayoutCachedWidth =24390
                    LayoutCachedHeight =2550
                    ColumnStart =25
                    ColumnEnd =25
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =25260
                    Top =2280
                    Width =0
                    Height =270
                    Name ="EmptyCell171"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =25260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =28
                    ColumnEnd =28
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =11895
                    Top =2280
                    Width =675
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label210"
                    Caption ="Days"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11895
                    LayoutCachedTop =2280
                    LayoutCachedWidth =12570
                    LayoutCachedHeight =2550
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Width =5040
                    Height =615
                    FontSize =22
                    FontWeight =700
                    Name ="Label261"
                    Caption ="EXIT Options Dashboard"
                    FontName ="Segoe UI"
                    RightPadding =38
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9930
                    Top =2280
                    Width =900
                    Height =270
                    FontSize =9
                    Name ="UnitMake_Label"
                    Caption ="Type"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9930
                    LayoutCachedTop =2280
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =2550
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =10860
                    Top =2280
                    Width =1005
                    Height =270
                    FontSize =9
                    Name ="Label306"
                    Caption ="Date"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10860
                    LayoutCachedTop =2280
                    LayoutCachedWidth =11865
                    LayoutCachedHeight =2550
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9330
                    Top =2280
                    Width =600
                    Height =270
                    FontSize =9
                    Name ="Label88"
                    Caption ="Month"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9330
                    LayoutCachedTop =2280
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =2550
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =9330
                    Top =2085
                    Width =1454
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="Label307"
                    Caption ="       ATF Options"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9330
                    LayoutCachedTop =2085
                    LayoutCachedWidth =10784
                    LayoutCachedHeight =2310
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =12600
                    Top =2280
                    Width =945
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label341"
                    Caption ="Date"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12600
                    LayoutCachedTop =2280
                    LayoutCachedWidth =13545
                    LayoutCachedHeight =2550
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =1860
                    Top =2280
                    Width =2955
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label344"
                    Caption ="  Schedule   |   Group"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1860
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =2550
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =30
                    Top =2280
                    Width =1770
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label357"
                    Caption ="Client"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =2280
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =2550
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =25260
                    Top =2280
                    Width =0
                    Height =270
                    Name ="EmptyCell367"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =25260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =29
                    ColumnEnd =29
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =25260
                    Top =2280
                    Width =0
                    Height =270
                    Name ="EmptyCell372"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =25260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =30
                    ColumnEnd =30
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =25260
                    Top =2280
                    Width =0
                    Height =270
                    Name ="EmptyCell377"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =25260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =31
                    ColumnEnd =31
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7380
                    Top =1500
                    Width =1673
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="ResetForm"
                    Caption ="Reset / Clear Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7380
                    LayoutCachedTop =1500
                    LayoutCachedWidth =9053
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2524406
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =8435191
                    PressedColor =2524406
                    HoverForeColor =1643706
                    QuickStyle =35
                    QuickStyleMask =-373
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9488
                    Top =600
                    Width =4620
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="OpenExitOptions"
                    Caption ="Open EXIT Options: "
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Open EXIT Option Form with the Group shown on the button"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9488
                    LayoutCachedTop =600
                    LayoutCachedWidth =14108
                    LayoutCachedHeight =840
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =14270637
                    PressedColor =8210719
                    HoverForeColor =6108695
                    PressedForeThemeColorIndex =1
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16435
                    Top =420
                    Width =1079
                    Height =420
                    ColumnOrder =21
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    BackColor =14211288
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=Count([AssetID])"
                    Format ="Standard"

                    LayoutCachedLeft =16435
                    LayoutCachedTop =420
                    LayoutCachedWidth =17514
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =16410
                            Top =210
                            Width =1140
                            Height =660
                            FontSize =9
                            FontWeight =700
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label63"
                            Caption ="Units Shown"
                            LayoutCachedLeft =16410
                            LayoutCachedTop =210
                            LayoutCachedWidth =17550
                            LayoutCachedHeight =870
                            BackThemeColorIndex =6
                            BackShade =50.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =119
                    TextFontCharSet =0
                    IMESentenceMode =3
                    ColumnCount =3
                    Top =840
                    Width =1920
                    ColumnOrder =17
                    FontSize =9
                    TabIndex =7
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT clientID, Client, Type FROM ClientMasterLeasesTypeCount_Summed Where Type"
                        " = 'FA' or  Type = 'Both' Order by Client;"
                    ColumnWidths ="0;2016;0"
                    OnClick ="[Event Procedure]"
                    Tag ="MEdit"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedTop =840
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =2280
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =119
                    TextFontCharSet =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =2040
                    Top =840
                    Width =5205
                    ColumnOrder =18
                    FontSize =9
                    TabIndex =8
                    Name ="SchGrpList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT GroupID, [Schedule] & ' | ' & [UnitGroup] AS [Sch|Grp], MLOrig, Count(Ass"
                        "etID) AS Units From ExitOptionsGroup_Local WHERE clientID =  GROUP BY GroupID, M"
                        "LOrig, Schedule, UnitGroup Order By GroupID Desc;"
                    ColumnWidths ="0;3600;864;432"
                    OnClick ="[Event Procedure]"
                    Tag ="MEdit"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2040
                    LayoutCachedTop =840
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =2280
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2040
                    Top =600
                    Width =5205
                    ColumnOrder =19
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    ForeColor =-2147483615
                    Name ="SearchML"
                    OnChange ="[Event Procedure]"
                    Tag ="MEdit"
                    ControlTipText ="Search MasterLease - Schedule - Group"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2040
                    LayoutCachedTop =600
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Top =600
                    Width =1914
                    ColumnOrder =20
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    ForeColor =-2147483615
                    Name ="SearchClient"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Tag ="MEdit"
                    ControlTipText ="Search Clients"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedTop =600
                    LayoutCachedWidth =1914
                    LayoutCachedHeight =840
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =10350
                    Top =1500
                    Width =1673
                    Height =240
                    ColumnOrder =16
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="Toggle120Exch"
                    Caption ="ExchangeIT < 180 Days"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter Units 120 or less days when they can ExchangeIT (toggle)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10350
                    LayoutCachedTop =1500
                    LayoutCachedWidth =12023
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =12600
                    Top =1500
                    Width =1643
                    Height =240
                    ColumnOrder =15
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    Name ="Toggle120Eol"
                    Caption ="All Lease Exp"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter Units 120 days until EOL (toggle)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12600
                    LayoutCachedTop =1500
                    LayoutCachedWidth =14243
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =12570
                    Top =2085
                    Width =960
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label505"
                    Caption ="Expiration "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12570
                    LayoutCachedTop =2085
                    LayoutCachedWidth =13530
                    LayoutCachedHeight =2310
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =8100
                    Top =1275
                    Width =11490
                    Height =225
                    FontSize =9
                    ForeColor =7500402
                    Name ="LinkLineFA"
                    Caption ="┌───────────── Filters ────────────┬──────────────────────────┬────────────┬────"
                        "────┬─────┴──┬─────┬────────┬───────────────┐"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8100
                    LayoutCachedTop =1275
                    LayoutCachedWidth =19590
                    LayoutCachedHeight =1500
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =15405
                    Top =1080
                    Width =1650
                    Height =202
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="ToggleFindNone"
                    Caption ="None Selected"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Unit Without An Exit Options Selected"

                    LayoutCachedLeft =15405
                    LayoutCachedTop =1080
                    LayoutCachedWidth =17055
                    LayoutCachedHeight =1282
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =14303
                    Top =1500
                    Width =945
                    Height =240
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    Name ="ToggleFindExtd"
                    Caption ="Extended"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Units that were Extended (toggle)"

                    LayoutCachedLeft =14303
                    LayoutCachedTop =1500
                    LayoutCachedWidth =15248
                    LayoutCachedHeight =1740
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =16823
                    Top =1500
                    Width =450
                    Height =240
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    Name ="ToggleFindTerm"
                    Caption ="Term"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Units Selected To Go Off Lease (toggle)"

                    LayoutCachedLeft =16823
                    LayoutCachedTop =1500
                    LayoutCachedWidth =17273
                    LayoutCachedHeight =1740
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =17333
                    Top =1500
                    Width =1230
                    Height =240
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    Name ="ToggleFindExchanged"
                    Caption ="Were Exchanged"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Units that were Exchanged"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =17333
                    LayoutCachedTop =1500
                    LayoutCachedWidth =18563
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =18623
                    Top =1500
                    Width =1260
                    Height =240
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    Name ="ToggleFindNOTExchanged"
                    Caption ="Yet To Exchange"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Units That Were NOT Yet Exchanged"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =18623
                    LayoutCachedTop =1500
                    LayoutCachedWidth =19883
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =25260
                    Top =2280
                    Width =0
                    Height =270
                    Name ="EmptyCell524"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =25260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =27
                    ColumnEnd =27
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =7260
                    Top =2280
                    Width =2040
                    Height =270
                    FontSize =9
                    ForeColor =6108695
                    Name ="Label517"
                    Caption ="┌─ Add/Edit Unit─┐"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7260
                    LayoutCachedTop =2280
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =2550
                    ColumnStart =5
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =7380
                    Top =1980
                    Width =1673
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =25
                    Name ="ExitOptionsHistory"
                    Caption ="Group EXIT History"
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Open the the Group's history (using the EXIT Options form)"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7380
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9053
                    LayoutCachedHeight =2220
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =9211020
                    OldBorderStyle =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =12566463
                    PressedColor =9211020
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    Overlaps =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    Left =7320
                    Top =600
                    Width =660
                    Height =195
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =0
                    ForeColor =11830108
                    Name ="ClearAll"
                    Caption ="◄Clear"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the Searches and List Boxes"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =600
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =795
                    ThemeFontIndex =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =15225
                    Top =2280
                    Width =975
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label614"
                    Caption ="Date"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15225
                    LayoutCachedTop =2280
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =2550
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =16313
                    Top =1500
                    Width =450
                    Height =240
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    Name ="ToggleFindOLAP"
                    Caption ="OVLP"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="ExchangedIT Units that have OVerLaPing leases with the New unit (toggle)"

                    LayoutCachedLeft =16313
                    LayoutCachedTop =1500
                    LayoutCachedWidth =16763
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =15225
                    Top =2085
                    Width =975
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label632"
                    Caption ="Off Lease"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15225
                    LayoutCachedTop =2085
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =2310
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =15308
                    Top =1500
                    Width =945
                    Height =240
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="ToggleOffLease"
                    Caption ="Off Lease"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Units that are Off Lease (toggle)"

                    LayoutCachedLeft =15308
                    LayoutCachedTop =1500
                    LayoutCachedWidth =16253
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9330
                    Top =1890
                    Width =3211
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="Label640"
                    Caption ="┌───────── ExchangeIT ────────┐"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9330
                    LayoutCachedTop =1890
                    LayoutCachedWidth =12541
                    LayoutCachedHeight =2115
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10856
                    Top =2085
                    Width =1706
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="Label641"
                    Caption ="       1st Notice"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10856
                    LayoutCachedTop =2085
                    LayoutCachedWidth =12562
                    LayoutCachedHeight =2310
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =12570
                    Top =1890
                    Width =2670
                    Height =225
                    FontSize =8
                    ForeColor =6108695
                    Name ="Label642"
                    Caption =" ┌ Pri Lease ─────┐┌─── Projected ┐"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12570
                    LayoutCachedTop =1890
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =2115
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =16200
                    Top =2280
                    Width =195
                    Height =270
                    Name ="EmptyCell655"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =16200
                    LayoutCachedTop =2280
                    LayoutCachedWidth =16395
                    LayoutCachedHeight =2550
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =21735
                    Top =2280
                    Width =864
                    Height =270
                    FontSize =9
                    Name ="Label665"
                    Caption ="Est Delv"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =21735
                    LayoutCachedTop =2280
                    LayoutCachedWidth =22599
                    LayoutCachedHeight =2550
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =15225
                    Top =1890
                    Width =975
                    Height =225
                    FontSize =8
                    ForeColor =6108695
                    Name ="Label667"
                    Caption ="┌ Actual ┐"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15225
                    LayoutCachedTop =1890
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =2115
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =18675
                    Top =2040
                    Width =5760
                    Height =180
                    BorderColor =6108695
                    Name ="Box671"
                    LayoutCachedLeft =18675
                    LayoutCachedTop =2040
                    LayoutCachedWidth =24435
                    LayoutCachedHeight =2220
                End
                Begin Rectangle
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =127
                    Left =18615
                    Top =2220
                    Width =6048
                    Height =60
                    Name ="Box672"
                    LayoutCachedLeft =18615
                    LayoutCachedTop =2220
                    LayoutCachedWidth =24663
                    LayoutCachedHeight =2280
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =19635
                    Top =1920
                    Width =1710
                    Height =240
                    FontSize =8
                    BackColor =14211288
                    ForeColor =6108695
                    Name ="ExchGroupLabel"
                    Caption =" EXchangeIT Equipment"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19635
                    LayoutCachedTop =1920
                    LayoutCachedWidth =21345
                    LayoutCachedHeight =2160
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin OptionGroup
                    Visible = NotDefault
                    BackStyle =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =5100
                    Top =240
                    Width =2240
                    Height =300
                    ColumnOrder =7
                    BackColor =15527148
                    BorderColor =0
                    Name ="LeaseTypeToShow"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5100
                    LayoutCachedTop =240
                    LayoutCachedWidth =7340
                    LayoutCachedHeight =540
                    Begin
                        Begin OptionButton
                            Visible = NotDefault
                            OverlapFlags =87
                            Left =5160
                            Top =270
                            Width =725
                            OptionValue =1
                            Name ="CheckAll"
                            ControlTipText ="FA Units (A, B, ,TR, R or PS Status)"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =5160
                            LayoutCachedTop =270
                            LayoutCachedWidth =5885
                            LayoutCachedHeight =510
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            Begin
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =5390
                                    Top =270
                                    Width =660
                                    Height =180
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label686"
                                    Caption ="FA Lease"
                                    LayoutCachedLeft =5390
                                    LayoutCachedTop =270
                                    LayoutCachedWidth =6050
                                    LayoutCachedHeight =450
                                End
                            End
                        End
                        Begin OptionButton
                            Visible = NotDefault
                            OverlapFlags =87
                            Left =6225
                            Top =270
                            Width =725
                            TabIndex =1
                            OptionValue =2
                            Name ="CheckLeased"
                            ControlTipText ="'Client Owned' Units (R, T or TR Status)"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =6225
                            LayoutCachedTop =270
                            LayoutCachedWidth =6950
                            LayoutCachedHeight =510
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            Begin
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =6455
                                    Top =270
                                    Width =690
                                    Height =210
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label687"
                                    Caption ="Portfolio"
                                    LayoutCachedLeft =6455
                                    LayoutCachedTop =270
                                    LayoutCachedWidth =7145
                                    LayoutCachedHeight =480
                                End
                            End
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =5130
                    Top =75
                    Width =2025
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label684"
                    Caption ="Lease Type Shown"
                    LayoutCachedLeft =5130
                    LayoutCachedTop =75
                    LayoutCachedWidth =7155
                    LayoutCachedHeight =315
                    ThemeFontIndex =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextFontCharSet =0
                    IMESentenceMode =3
                    ListWidth =720
                    Left =21360
                    Top =1485
                    Width =720
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =11
                    BorderColor =8210719
                    Name ="EOLYear"
                    RowSourceType ="Value List"
                    RowSource ="\"2018\";\"2019\";\"2020\";\"2021\";\"2022\";\"2023\";\"2024\";\"2025\""
                    ColumnWidths ="720"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Select a Year to Filer By"
                    Format ="General Number"

                    LayoutCachedLeft =21360
                    LayoutCachedTop =1485
                    LayoutCachedWidth =22080
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =21360
                            Top =1260
                            Width =2610
                            Height =240
                            FontSize =9
                            BorderColor =0
                            Name ="Label340"
                            Caption ="Year ──────┬────┬────┬───┐"
                            Tag ="FAOnly"
                            LayoutCachedLeft =21360
                            LayoutCachedTop =1260
                            LayoutCachedWidth =23970
                            LayoutCachedHeight =1500
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =22140
                    Top =1515
                    Width =465
                    Height =225
                    ColumnOrder =2
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    Name ="ToggleQTR1"
                    Caption ="Q 1"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter By Year - Quarter 1"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =22140
                    LayoutCachedTop =1515
                    LayoutCachedWidth =22605
                    LayoutCachedHeight =1740
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =5
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =22620
                    Top =1515
                    Width =465
                    Height =225
                    ColumnOrder =3
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
                    Name ="ToggleQTR2"
                    Caption ="Q 2"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter By Year - Quarter 2"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =22620
                    LayoutCachedTop =1515
                    LayoutCachedWidth =23085
                    LayoutCachedHeight =1740
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =5
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =23100
                    Top =1515
                    Width =465
                    Height =225
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    Name ="ToggleQTR3"
                    Caption ="Q 3"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter By Year - Quarter 2"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23100
                    LayoutCachedTop =1515
                    LayoutCachedWidth =23565
                    LayoutCachedHeight =1740
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =5
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =23580
                    Top =1515
                    Width =465
                    Height =225
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
                    Name ="ToggleQTR4"
                    Caption ="Q 4"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Filter By Year - Quarter 4"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23580
                    LayoutCachedTop =1515
                    LayoutCachedWidth =24045
                    LayoutCachedHeight =1740
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =5
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin OptionGroup
                    BackStyle =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =21360
                    Top =945
                    Width =2240
                    Height =300
                    ColumnOrder =6
                    TabIndex =9
                    BackColor =15527148
                    BorderColor =0
                    Name ="EndTypeToShow"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =21360
                    LayoutCachedTop =945
                    LayoutCachedWidth =23600
                    LayoutCachedHeight =1245
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            Left =21420
                            Top =975
                            Width =725
                            OptionValue =1
                            Name ="Option708"
                            Tag ="FAOnly"
                            ControlTipText ="Show when Units are schedued to end their lease"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =21420
                            LayoutCachedTop =975
                            LayoutCachedWidth =22145
                            LayoutCachedHeight =1215
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =21650
                                    Top =975
                                    Width =660
                                    Height =180
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label709"
                                    Caption ="Expiry"
                                    Tag ="FAOnly"
                                    LayoutCachedLeft =21650
                                    LayoutCachedTop =975
                                    LayoutCachedWidth =22310
                                    LayoutCachedHeight =1155
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =22485
                            Top =975
                            Width =725
                            OptionValue =2
                            Name ="Option710"
                            Tag ="FAOnly"
                            ControlTipText ="Show when Units Can be 1st Exchanged"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =22485
                            LayoutCachedTop =975
                            LayoutCachedWidth =23210
                            LayoutCachedHeight =1215
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =22710
                                    Top =975
                                    Width =855
                                    Height =210
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label711"
                                    Caption ="EXchangeIT"
                                    Tag ="FAOnly"
                                    LayoutCachedLeft =22710
                                    LayoutCachedTop =975
                                    LayoutCachedWidth =23565
                                    LayoutCachedHeight =1185
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21360
                    Top =780
                    Width =2085
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =0
                    ForeColor =6108695
                    Name ="Label712"
                    Caption ="Show When Units"
                    Tag ="FAOnly"
                    LayoutCachedLeft =21360
                    LayoutCachedTop =780
                    LayoutCachedWidth =23445
                    LayoutCachedHeight =1020
                    ThemeFontIndex =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =14168
                    Top =600
                    Width =1380
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="ExportToExcel"
                    Caption =" Quick Export"
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Export what is displayed to Excel"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
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

                    LayoutCachedLeft =14168
                    LayoutCachedTop =600
                    LayoutCachedWidth =15548
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =5
                    Alignment =1
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    PressedColor =9211020
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8010
                    Top =600
                    Width =1418
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="SyncGetLatestData"
                    Caption =" Sync|Get Data"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0x727272ff727272ff727272ff727272ff00000000727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ffffffffffffffffff727272ff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ffffffffffffffffff727272ff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000b3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8010
                    LayoutCachedTop =600
                    LayoutCachedWidth =9428
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2524406
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =8435191
                    PressedColor =2524406
                    HoverForeColor =1643706
                    QuickStyle =35
                    QuickStyleMask =-373
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =14280
                    Top =2280
                    Width =945
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label725"
                    Caption ="Date"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14280
                    LayoutCachedTop =2280
                    LayoutCachedWidth =15225
                    LayoutCachedHeight =2550
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14280
                    Top =2085
                    Width =945
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label727"
                    Caption ="Off Lease"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14280
                    LayoutCachedTop =2085
                    LayoutCachedWidth =15225
                    LayoutCachedHeight =2310
                    ColumnStart =7
                    ColumnEnd =10
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13605
                    Top =2085
                    Width =645
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label729"
                    Caption ="Days"
                    GroupTable =6
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13605
                    LayoutCachedTop =2085
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =2325
                    LayoutGroup =4
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =24390
                    Top =2280
                    Width =870
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label735"
                    Caption ="In Service"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =24390
                    LayoutCachedTop =2280
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =2550
                    ColumnStart =26
                    ColumnEnd =26
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =22650
                    Top =2280
                    Width =870
                    Height =270
                    FontSize =9
                    Name ="Label737"
                    Caption ="Act Delv"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =22650
                    LayoutCachedTop =2280
                    LayoutCachedWidth =23520
                    LayoutCachedHeight =2550
                    ColumnStart =24
                    ColumnEnd =24
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =24660
                    Top =300
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =26
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =24660
                    LayoutCachedTop =300
                    LayoutCachedWidth =25215
                    LayoutCachedHeight =528
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =0
                    PictureType =1
                    Left =25227
                    Top =300
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =27
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =25227
                    LayoutCachedTop =300
                    LayoutCachedWidth =25782
                    LayoutCachedHeight =528
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
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =18060
                    Top =600
                    Width =1380
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =28
                    Name ="ExportSendEmail"
                    Caption =" Expire Email"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c8760009c7750036c7730066c6720096 ,
                        0xc57100c3c56f00f0c46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800aec97700dbc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ffc67200ffc67200ffc67200ffc67200ff ,
                        0xc67200f6c672008dc97800ffc97700ffc87600ffe4bc84fffcf7f0fffcf7f0ff ,
                        0xe4bc87ffc56f00ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xfdfaf6ffc67200f6c97800ffc97700ffdfae69ffffffffffeaca9fffefd7b7ff ,
                        0xffffffffdca966ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fff5e5cfffeacb9fffc77300ffc67200ff ,
                        0xf3e3ccfff3e2ccffc46e00ffc36d00fffffffffff8eddeffffffffffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fffdfaf6ffddab63ffc77300ffc67200ff ,
                        0xe8c699fffdfaf6ffc46e00ffc36d00ffe8c799ffc67200ffe1b67bffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fffcf7f0ffdca960ffc77300ffc67200ff ,
                        0xe8c699fffefcf9ffc46e00ffc36d00ffc77506ffe1b67bffc87709ffe1b478ff ,
                        0xffffffffc67200ffc97800ffc97700fff3e0c6ffebcea5ffc77300ffc67200ff ,
                        0xf5e8d5fff4e4cfffc46e00ffc36d00ffedd2aefffffffffff2dec3ffc9790cff ,
                        0xe0b375ffc67200ffc97800ffc97700ffdca85dffffffffffeccfa8ffefd7b7ff ,
                        0xffffffffdfaf72ffc46e00ffc36d00fffffffffffffffffffffffffff2dec3ff ,
                        0xc9790cffc67200ffc97800ffc97700ffc87600ffe1b475fffaf2e7fffcf8f3ff ,
                        0xe4bc87ffc67103ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xf0dbbfffc67200f6c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ffc67200ffc67200ffc67200ffc67200ff ,
                        0xc67200f6c672008dc97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800aec97700dbc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c8760009c7750033c7730066c6720093 ,
                        0xc57100c0c56f00edc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =18060
                    LayoutCachedTop =600
                    LayoutCachedWidth =19440
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =19560
                    Top =600
                    Width =1380
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =29
                    Name ="SUFEmail"
                    Caption =" SUF  Email"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c8760009c7750036c7730066c6720096 ,
                        0xc57100c3c56f00f0c46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800aec97700dbc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ffc67200ffc67200ffc67200ffc67200ff ,
                        0xc67200f6c672008dc97800ffc97700ffc87600ffe4bc84fffcf7f0fffcf7f0ff ,
                        0xe4bc87ffc56f00ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xfdfaf6ffc67200f6c97800ffc97700ffdfae69ffffffffffeaca9fffefd7b7ff ,
                        0xffffffffdca966ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fff5e5cfffeacb9fffc77300ffc67200ff ,
                        0xf3e3ccfff3e2ccffc46e00ffc36d00fffffffffff8eddeffffffffffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fffdfaf6ffddab63ffc77300ffc67200ff ,
                        0xe8c699fffdfaf6ffc46e00ffc36d00ffe8c799ffc67200ffe1b67bffffffffff ,
                        0xffffffffc67200ffc97800ffc97700fffcf7f0ffdca960ffc77300ffc67200ff ,
                        0xe8c699fffefcf9ffc46e00ffc36d00ffc77506ffe1b67bffc87709ffe1b478ff ,
                        0xffffffffc67200ffc97800ffc97700fff3e0c6ffebcea5ffc77300ffc67200ff ,
                        0xf5e8d5fff4e4cfffc46e00ffc36d00ffedd2aefffffffffff2dec3ffc9790cff ,
                        0xe0b375ffc67200ffc97800ffc97700ffdca85dffffffffffeccfa8ffefd7b7ff ,
                        0xffffffffdfaf72ffc46e00ffc36d00fffffffffffffffffffffffffff2dec3ff ,
                        0xc9790cffc67200ffc97800ffc97700ffc87600ffe1b475fffaf2e7fffcf8f3ff ,
                        0xe4bc87ffc67103ffc46e00ffc36d00ffffffffffffffffffffffffffffffffff ,
                        0xf0dbbfffc67200f6c97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ffc67200ffc67200ffc67200ffc67200ff ,
                        0xc67200f6c672008dc97800ffc97700ffc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000c97800aec97700dbc87600ffc77500ffc77300ffc67200ff ,
                        0xc57100ffc56f00ffc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c8760009c7750033c7730066c6720093 ,
                        0xc57100c0c56f00edc46e00ffc36d00ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =19560
                    LayoutCachedTop =600
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
            End
        End
        Begin Section
            Height =359
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13605
                    Width =645
                    Height =270
                    ColumnWidth =1410
                    ColumnOrder =15
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    Name ="DaysEol"
                    ControlSource ="=IIf([DaysToEOL]<0 Or IsDate([ActualOffLease]),\"--\",[DaysToEOL])"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13605
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4875
                    Width =780
                    Height =270
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4875
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25260
                    Width =0
                    Height =270
                    ColumnWidth =3000
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =29
                    Name ="GroupId"
                    ControlSource ="GroupId"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =25260
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =27
                    ColumnEnd =27
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6660
                    Width =540
                    Height =270
                    ColumnWidth =840
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BorderColor =14211288
                    Name ="uStatus"
                    ControlSource ="uStatus"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5715
                    Width =885
                    Height =270
                    ColumnWidth =1500
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =6
                    Name ="UnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5715
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18165
                    Width =2664
                    Height =270
                    FontSize =9
                    TabIndex =23
                    BorderColor =14211288
                    ForeColor =8210719
                    Name ="NewSchGrp"
                    ControlSource ="=IIf(IsNull([newSchedule]),Null,[NewClient] & \" | \" & [NewSchedule] & \" | \" "
                        "& [NewGroup])"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18165
                    LayoutCachedWidth =20829
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20895
                    Width =780
                    Height =270
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =24
                    Name ="NewUnit"
                    ControlSource ="NewAssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =20895
                    LayoutCachedWidth =21675
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25260
                    Width =0
                    Height =270
                    FontSize =9
                    TabIndex =30
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =25260
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =28
                    ColumnEnd =28
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23520
                    Width =870
                    Height =270
                    FontSize =9
                    TabIndex =27
                    Name ="LastModified"
                    ControlSource ="nAccepted"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =23520
                    LayoutCachedWidth =24390
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =25
                    ColumnEnd =25
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =7260
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =8
                    Name ="ModifyExch"
                    Caption ="EXCH"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set ExchangeIT Option for this unit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7260
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8872517
                    BorderColor =15527148
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11895
                    Width =675
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    Name ="ExchDays"
                    ControlSource ="=IIf([DaysToExch]>=0 And [EXCH]=\"----\",[DaysToExch],\"-\")"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11895
                    LayoutCachedWidth =12570
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =8310
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =10
                    Name ="ModifyTERM"
                    Caption ="TERM"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Set Terminate Option for this unit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8310
                    LayoutCachedWidth =8775
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =790406
                    BorderColor =15527148
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =790406
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =7785
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =9
                    Name ="ModifyEXTD"
                    Caption ="EXTD"
                    OnClick ="[Event Procedure]"
                    Tag ="FAOnly"
                    ControlTipText ="Set Extend Option for this unit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7785
                    LayoutCachedWidth =8250
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =0
                    BorderColor =15527148
                    HoverColor =9211020
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9930
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =13
                    Name ="UnitMake"
                    ControlSource ="ExchTyp"
                    StatusBarText ="unit make"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9930
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9330
                    Width =600
                    Height =270
                    FontSize =9
                    TabIndex =12
                    Name ="ExchangeITMonth2"
                    ControlSource ="=IIf([ExchangeITMonth]=0,\"--\",[ExchangeITMonth])"
                    StatusBarText ="unit model"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9330
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10860
                    Width =1005
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="ExchDate"
                    ControlSource ="ExchDate"
                    Format ="Short Date"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10860
                    LayoutCachedWidth =11865
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Width =945
                    Height =270
                    FontSize =9
                    TabIndex =16
                    Name ="EOLPrimary"
                    ControlSource ="EOLPrimary"
                    Format ="Short Date"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12600
                    LayoutCachedWidth =13545
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1860
                    Width =2955
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="ScheduleGroup"
                    ControlSource ="=[ReplacedSchNo] & \" | \" & [ReplacedGroupNo] & \" | \" & [UnitDesc]"
                    StatusBarText ="unit model"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1860
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Width =1770
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="Client"
                    ControlSource ="Client"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25260
                    Width =0
                    Height =270
                    FontSize =9
                    TabIndex =31
                    Name ="ClientID"
                    ControlSource ="ClientID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =25260
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =29
                    ColumnEnd =29
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25260
                    Width =0
                    Height =270
                    FontSize =9
                    TabIndex =32
                    Name ="NewSchID"
                    ControlSource ="ReplacedByScheduleID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =25260
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =30
                    ColumnEnd =30
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25260
                    Width =0
                    Height =270
                    FontSize =9
                    TabIndex =33
                    Name ="NewGroupID"
                    ControlSource ="ReplacedByGroupID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =25260
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =31
                    ColumnEnd =31
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Top =300
                    Width =26424
                    Height =59
                    TabIndex =2
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2
                    VerticalAnchor =1

                    LayoutCachedTop =300
                    LayoutCachedWidth =26424
                    LayoutCachedHeight =359
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22800
                    Top =288
                    Width =825
                    Height =60
                    FontSize =9
                    Name ="OldGroupID"
                    ControlSource ="OldGroupID"
                    TopPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =22800
                    LayoutCachedTop =288
                    LayoutCachedWidth =23625
                    LayoutCachedHeight =348
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23760
                    Top =288
                    Width =825
                    Height =60
                    FontSize =9
                    TabIndex =1
                    Name ="OldSchID"
                    ControlSource ="OldSchID"
                    TopPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =23760
                    LayoutCachedTop =288
                    LayoutCachedWidth =24585
                    LayoutCachedHeight =348
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =8835
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =11
                    Name ="ModifyHIST"
                    Caption ="HIST"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Terminate Option for this unit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8835
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =9211020
                    BorderColor =15527148
                    HoverColor =12566463
                    PressedColor =790406
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16395
                    Width =525
                    Height =270
                    FontSize =9
                    TabIndex =20
                    Name ="EXCH"
                    ControlSource ="=Left([Selected],4)"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16395
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16950
                    Width =510
                    Height =270
                    FontSize =9
                    TabIndex =21
                    Name ="EXTD"
                    ControlSource ="=Mid([Selected],7,4)"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16950
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17520
                    Width =585
                    Height =270
                    FontSize =9
                    TabIndex =22
                    Name ="Text599"
                    ControlSource ="=Right([Selected],4)"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17520
                    LayoutCachedWidth =18105
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15225
                    Width =975
                    Height =270
                    FontSize =9
                    TabIndex =19
                    Name ="ActualOffLease"
                    ControlSource ="ActualOffLease"
                    Format ="Short Date"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15225
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21735
                    Width =864
                    Height =270
                    FontSize =9
                    TabIndex =25
                    Name ="Text659"
                    ControlSource ="=[nEstDelv]"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =21735
                    LayoutCachedWidth =22599
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14280
                    Width =945
                    Height =270
                    FontSize =9
                    TabIndex =18
                    Name ="EOLDate"
                    ControlSource ="EOLDate"
                    Format ="Short Date"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14280
                    LayoutCachedWidth =15225
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =16200
                    Width =195
                    Height =270
                    Name ="EmptyCell730"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =16200
                    LayoutCachedWidth =16395
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =24390
                    Width =870
                    Height =270
                    FontSize =9
                    TabIndex =28
                    Name ="Text733"
                    ControlSource ="nInServ"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =24390
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =26
                    ColumnEnd =26
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22650
                    Width =870
                    Height =270
                    FontSize =9
                    TabIndex =26
                    Name ="Text743"
                    ControlSource ="[nActualDelv]"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =22650
                    LayoutCachedWidth =23520
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =24
                    ColumnEnd =24
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GroupTable =3
                End
            End
        End
        Begin FormFooter
            Height =210
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    FontSize =8
                    BorderColor =12632256
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1020
                    Width =765
                    Height =210
                    FontSize =8
                    TabIndex =1
                    BorderColor =12632256
                    Name ="ScheduleIDx"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1920
                    Width =765
                    Height =210
                    FontSize =8
                    TabIndex =2
                    BorderColor =12632256
                    Name ="GroupIDx"

                    LayoutCachedLeft =1920
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =210
                End
            End
        End
    End
End
CodeBehindForm
' See "ExitOptionsManager.cls"
