Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21443
    DatasheetFontHeight =11
    ItemSuffix =365
    Left =2745
    Top =4890
    Right =23595
    Bottom =17580
    Filter ="([ReLeaseUnitsSub].[Client] In (\"BB8\",\"FA RMKT\",\"Transco\"))"
    ShortcutMenuBar ="SortClip"
    OrderBy ="[ReLeaseUnitsSub].[Vin], [ReLeaseUnitsSub].[AssetID]"
    RecSrcDt = Begin
        0xac07e07e31f1e540
    End
    RecordSource ="SELECT * FROM vw_ReLeasedUnits_LastTitleHistory  WHERE Not AssetID is null  ORDE"
        "R BY VIN, AssetID"
    Caption ="UnitsIn_sub"
    OnCurrent ="[Event Procedure]"
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
            Height =375
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =165
                    Width =20535
                    Height =210
                    BorderColor =15527148
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =165
                    LayoutCachedWidth =20535
                    LayoutCachedHeight =375
                    TabIndex =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =19065
                    Top =30
                    Width =600
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitFinalCost_Label"
                    Caption ="Return"
                    GroupTable =6
                    LayoutCachedLeft =19065
                    LayoutCachedTop =30
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =300
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4515
                    Top =30
                    Width =1005
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="♦AssetID"
                    GroupTable =6
                    LayoutCachedLeft =4515
                    LayoutCachedTop =30
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =300
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    Left =120
                    Top =30
                    Width =1875
                    Height =270
                    FontSize =9
                    Name ="ULoc_Label"
                    Caption ="Client"
                    GroupTable =6
                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =300
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =12195
                    Top =30
                    Width =1005
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitEstDelDate_Label"
                    Caption ="Title Date"
                    GroupTable =6
                    LayoutCachedLeft =12195
                    LayoutCachedTop =30
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =300
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =5580
                    Top =30
                    Width =1950
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitVIN_Label"
                    Caption ="VIN"
                    GroupTable =6
                    LayoutCachedLeft =5580
                    LayoutCachedTop =30
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =300
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =7590
                    Top =30
                    Width =2835
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitYr_Label"
                    Caption ="Unit Description"
                    GroupTable =6
                    LayoutCachedLeft =7590
                    LayoutCachedTop =30
                    LayoutCachedWidth =10425
                    LayoutCachedHeight =300
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    Left =13260
                    Top =30
                    Width =3060
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label86"
                    Caption ="Reason"
                    GroupTable =6
                    LayoutCachedLeft =13260
                    LayoutCachedTop =30
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =300
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =19725
                    Top =30
                    Width =540
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label106"
                    Caption ="Type"
                    GroupTable =6
                    LayoutCachedLeft =19725
                    LayoutCachedTop =30
                    LayoutCachedWidth =20265
                    LayoutCachedHeight =300
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =2760
                    Top =30
                    Width =900
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="unitinserv_Label"
                    Caption ="Schedule"
                    GroupTable =6
                    LayoutCachedLeft =2760
                    LayoutCachedTop =30
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =300
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =3720
                    Top =30
                    Width =735
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitCOASent_Label"
                    Caption ="Group"
                    GroupTable =6
                    LayoutCachedLeft =3720
                    LayoutCachedTop =30
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1080
                    Width =765
                    Height =210
                    FontSize =8
                    BorderColor =12632256
                    Name ="txtCurrent"

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =1845
                    LayoutCachedHeight =210
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =11025
                    Top =30
                    Width =1110
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label299"
                    Caption ="Assignee"
                    GroupTable =6
                    LayoutCachedLeft =11025
                    LayoutCachedTop =30
                    LayoutCachedWidth =12135
                    LayoutCachedHeight =300
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =20325
                    Top =30
                    Width =285
                    Height =270
                    Name ="EmptyCell307"
                    GroupTable =6
                    LayoutCachedLeft =20325
                    LayoutCachedTop =30
                    LayoutCachedWidth =20610
                    LayoutCachedHeight =300
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =16380
                    Top =30
                    Width =1950
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label344"
                    Caption ="Title Number"
                    GroupTable =6
                    LayoutCachedLeft =16380
                    LayoutCachedTop =30
                    LayoutCachedWidth =18330
                    LayoutCachedHeight =300
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =18390
                    Top =30
                    Width =615
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label346"
                    Caption ="State"
                    GroupTable =6
                    LayoutCachedLeft =18390
                    LayoutCachedTop =30
                    LayoutCachedWidth =19005
                    LayoutCachedHeight =300
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10485
                    Top =30
                    Width =480
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label355"
                    Caption ="STS"
                    GroupTable =6
                    LayoutCachedLeft =10485
                    LayoutCachedTop =30
                    LayoutCachedWidth =10965
                    LayoutCachedHeight =300
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    Left =2055
                    Top =30
                    Width =645
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label363"
                    Caption ="BDEX"
                    GroupTable =6
                    LayoutCachedLeft =2055
                    LayoutCachedTop =30
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
            End
        End
        Begin Section
            Height =344
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Width =645
                    Height =255
                    FontSize =9
                    Name ="UnitTitleLogID"
                    ControlSource ="UnitTitleLogID"

                    LayoutCachedWidth =645
                    LayoutCachedHeight =255
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13260
                    Top =30
                    Width =3060
                    Height =255
                    FontSize =9
                    TabIndex =11
                    BorderColor =12765388
                    Name ="lastReason"
                    ControlSource ="LastReason"
                    GroupTable =6

                    LayoutCachedLeft =13260
                    LayoutCachedTop =30
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4515
                    Top =30
                    Width =1005
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="AssetID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =4515
                    LayoutCachedTop =30
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12195
                    Top =30
                    Width =1005
                    Height =255
                    ColumnWidth =2700
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =10
                    BorderColor =12765388
                    Name ="LastDate"
                    ControlSource ="LastDate"
                    Format ="mm/dd/yy"
                    GroupTable =6

                    LayoutCachedLeft =12195
                    LayoutCachedTop =30
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5580
                    Top =30
                    Width =1950
                    Height =255
                    ColumnWidth =2100
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =6
                    Name ="UnitVIN"
                    ControlSource ="Vin"
                    StatusBarText ="VIN number must be unique in table"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =5580
                    LayoutCachedTop =30
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7590
                    Top =30
                    Width =2835
                    Height =255
                    ColumnWidth =990
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =7
                    Name ="UnitYr"
                    ControlSource ="UnitDesc"
                    GroupTable =6

                    LayoutCachedLeft =7590
                    LayoutCachedTop =30
                    LayoutCachedWidth =10425
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3720
                    Top =30
                    Width =735
                    Height =255
                    ColumnWidth =1290
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =4
                    BorderColor =12765388
                    Name ="UnitGroup"
                    ControlSource ="UnitGroup"
                    Format ="mm/dd/yy"
                    GroupTable =6

                    LayoutCachedLeft =3720
                    LayoutCachedTop =30
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19065
                    Top =30
                    Width =600
                    Height =255
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =14
                    BorderColor =12765388
                    Name ="ExpectedReturn"
                    ControlSource ="ExpectedReturn"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =6

                    LayoutCachedLeft =19065
                    LayoutCachedTop =30
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =30
                    Width =1875
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =1
                    Name ="ULoc"
                    ControlSource ="Client"
                    GroupTable =6

                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =30
                    Width =900
                    Height =255
                    ColumnWidth =1275
                    FontSize =9
                    TabIndex =3
                    BorderColor =12765388
                    Name ="Schedule"
                    ControlSource ="Schedule"
                    GroupTable =6

                    LayoutCachedLeft =2760
                    LayoutCachedTop =30
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =29
                    Top =285
                    Width =20160
                    Height =59
                    FontSize =10
                    TabIndex =16
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560049004e005d003d005b00740078007400430075007200720065006e00 ,
                        0x74005d0000000000
                    End

                    LayoutCachedLeft =29
                    LayoutCachedTop =285
                    LayoutCachedWidth =20189
                    LayoutCachedHeight =344
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00120000005b00 ,
                        0x560049004e005d003d005b00740078007400430075007200720065006e007400 ,
                        0x5d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11025
                    Top =30
                    Width =1110
                    Height =255
                    FontSize =9
                    TabIndex =9
                    Name ="Assignee"
                    ControlSource ="Assignee"
                    GroupTable =6

                    LayoutCachedLeft =11025
                    LayoutCachedTop =30
                    LayoutCachedWidth =12135
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19725
                    Top =30
                    Width =540
                    Height =255
                    FontSize =9
                    TabIndex =15
                    BorderColor =12765388
                    Name ="InOut"
                    ControlSource ="InOut"
                    ConditionalFormat = Begin
                        0x010000007a000000020000000000000002000000000000000500000001000000 ,
                        0x00000000cddcaf000000000002000000060000000c0000000100000000000000 ,
                        0xffcc660000000000000000000000000000000000000000000000000000000000 ,
                        0x220049004e0022000000000022004f0075007400220000000000
                    End
                    GroupTable =6

                    LayoutCachedLeft =19725
                    LayoutCachedTop =30
                    LayoutCachedWidth =20265
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000cddcaf00040000002200 ,
                        0x49004e0022000000000000000000000000000000000000000000000000000002 ,
                        0x0000000100000000000000ffcc66000500000022004f00750074002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16380
                    Top =30
                    Width =1950
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="Text333"
                    ControlSource ="TitleNumber"
                    GroupTable =6

                    LayoutCachedLeft =16380
                    LayoutCachedTop =30
                    LayoutCachedWidth =18330
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18390
                    Top =30
                    Width =615
                    Height =255
                    FontSize =9
                    TabIndex =13
                    Name ="Text335"
                    ControlSource ="TitleSt"
                    GroupTable =6

                    LayoutCachedLeft =18390
                    LayoutCachedTop =30
                    LayoutCachedWidth =19005
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =20325
                    Top =30
                    Width =285
                    Height =255
                    Name ="EmptyCell348"
                    GroupTable =6
                    LayoutCachedLeft =20325
                    LayoutCachedTop =30
                    LayoutCachedWidth =20610
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10485
                    Top =30
                    Width =480
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="UnitStatus"
                    ControlSource ="Status"
                    GroupTable =6

                    LayoutCachedLeft =10485
                    LayoutCachedTop =30
                    LayoutCachedWidth =10965
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2055
                    Top =30
                    Width =645
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =12765388
                    Name ="Text357"
                    ControlSource ="BDEX"
                    Format ="mm/dd/yy"
                    GroupTable =6

                    LayoutCachedLeft =2055
                    LayoutCachedTop =30
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
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
' See "ReLeaseUnitsSub.cls"
