Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18030
    DatasheetFontHeight =11
    ItemSuffix =57
    Left =8460
    Top =5640
    Right =27030
    Bottom =9705
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x7d80657eef7de440
    End
    RecordSource ="SELECT Bills.*, MstrLease.MLNo, Schedule.SchFSL FROM (MstrLease INNER JOIN Sched"
        "ule ON MstrLease.MLID = Schedule.SchMLRef) INNER JOIN Bills ON Schedule.SchID = "
        "Bills.BillSchRef ORDER BY Bills.BillID DESC , Bills.Billdt DESC;"
    OnClose ="[Event Procedure]"
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
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =720
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =540
                    Width =17784
                    Height =180
                    Name ="Image32"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =17784
                    LayoutCachedHeight =720
                    TabIndex =2
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =36
                    Top =360
                    Width =1956
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label17"
                    Caption ="Bill Type"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =36
                    LayoutCachedTop =360
                    LayoutCachedWidth =1992
                    LayoutCachedHeight =675
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =2064
                    Top =360
                    Width =1560
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label18"
                    Caption ="Invoice Number"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =2064
                    LayoutCachedTop =360
                    LayoutCachedWidth =3624
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =3696
                    Top =360
                    Width =1296
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label20"
                    Caption ="Bill Date"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =3696
                    LayoutCachedTop =360
                    LayoutCachedWidth =4992
                    LayoutCachedHeight =675
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =5664
                    Top =360
                    Width =1296
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label21"
                    Caption ="Bill From"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =5664
                    LayoutCachedTop =360
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =675
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =7032
                    Top =360
                    Width =1296
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label22"
                    Caption ="Bill To"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =7032
                    LayoutCachedTop =360
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =675
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =8400
                    Top =360
                    Width =1296
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label23"
                    Caption ="Bill Due"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =8400
                    LayoutCachedTop =360
                    LayoutCachedWidth =9696
                    LayoutCachedHeight =675
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10236
                    Top =360
                    Width =1296
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label24"
                    Caption ="Bill Sent"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =10236
                    LayoutCachedTop =360
                    LayoutCachedWidth =11532
                    LayoutCachedHeight =675
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11604
                    Top =360
                    Width =2916
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label25"
                    Caption ="Bill File Link"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =11604
                    LayoutCachedTop =360
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =675
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14592
                    Top =360
                    Width =648
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label30"
                    Caption ="Bill Id"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =14592
                    LayoutCachedTop =360
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =675
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =5064
                    Top =360
                    Width =528
                    Height =315
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label34"
                    Caption ="FSL?"
                    GroupTable =1
                    LayoutCachedLeft =5064
                    LayoutCachedTop =360
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =675
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =119
                    TextAlign =1
                    Left =120
                    Width =4260
                    Height =360
                    FontSize =16
                    FontWeight =700
                    BackColor =15707392
                    Name ="Label37"
                    Caption ="Add A New Bill"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9600
                    Top =240
                    Width =480
                    Height =420
                    FontSize =8
                    BackColor =15707392
                    Name ="Label41"
                    Caption ="Due  \015\012on\015\012Rec"
                    LayoutCachedLeft =9600
                    LayoutCachedTop =240
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =16980
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =16980
                    LayoutCachedTop =120
                    LayoutCachedWidth =17535
                    LayoutCachedHeight =348
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
                    OverlapFlags =215
                    Left =16380
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16380
                    LayoutCachedTop =120
                    LayoutCachedWidth =16935
                    LayoutCachedHeight =348
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
                Begin EmptyCell
                    Left =9768
                    Top =360
                    Width =396
                    Height =315
                    Name ="EmptyCell47"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =9768
                    LayoutCachedTop =360
                    LayoutCachedWidth =10164
                    LayoutCachedHeight =675
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15312
                    Top =360
                    Width =348
                    Height =315
                    Name ="EmptyCell55"
                    GroupTable =1
                    RightPadding =38
                    GridlineColor =0
                    LayoutCachedLeft =15312
                    LayoutCachedTop =360
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =675
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2064
                    Top =48
                    Width =1560
                    Height =300
                    TabIndex =4
                    Name ="BillNo"
                    ControlSource ="BillNo"
                    StatusBarText ="from Quickbooks or bank if type is BankBill"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Bill Number is required"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2064
                    LayoutCachedTop =48
                    LayoutCachedWidth =3624
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3696
                    Top =48
                    Width =1296
                    Height =300
                    TabIndex =5
                    Name ="Billdt"
                    ControlSource ="Billdt"
                    Format ="mm/dd/yy"
                    StatusBarText ="date bill generated"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3696
                    LayoutCachedTop =48
                    LayoutCachedWidth =4992
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5664
                    Top =48
                    Width =1296
                    Height =300
                    TabIndex =7
                    Name ="Billfrom"
                    ControlSource ="Billfrom"
                    Format ="mm/dd/yy"
                    StatusBarText ="start of billing period"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5664
                    LayoutCachedTop =48
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7032
                    Top =48
                    Width =1296
                    Height =300
                    TabIndex =8
                    Name ="Billto"
                    ControlSource ="Billto"
                    Format ="mm/dd/yy"
                    StatusBarText ="end of billing period"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7032
                    LayoutCachedTop =48
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8400
                    Top =48
                    Width =1296
                    Height =300
                    TabIndex =9
                    Name ="BillDueDt"
                    ControlSource ="BillDueDt"
                    StatusBarText ="date pmt due (defaults to start of billing period)"
                    ControlTipText ="date pmt due (defaults to start of billing period)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8400
                    LayoutCachedTop =48
                    LayoutCachedWidth =9696
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10236
                    Top =48
                    Width =1296
                    Height =300
                    TabIndex =11
                    Name ="BillSent"
                    ControlSource ="BillSent"
                    StatusBarText ="date bill sent to Lessee"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10236
                    LayoutCachedTop =48
                    LayoutCachedWidth =11532
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11604
                    Top =48
                    Width =2916
                    Height =300
                    TabIndex =12
                    ForeColor =1279872587
                    Name ="BillLink"
                    ControlSource ="BillLink"
                    StatusBarText ="hyperlink to scan of billing package."
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =11604
                    LayoutCachedTop =48
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =36
                    Top =48
                    Width =1956
                    Height =300
                    TabIndex =3
                    Name ="BillType"
                    ControlSource ="BillType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_BillType.BillType, Lookup_BillType.BillTypeID FROM Lookup_BillType"
                        " WHERE (((Lookup_BillType.BillTypeID)>0) AND ((Lookup_BillType.Obsolete)=False))"
                        ";"
                    ColumnWidths ="1440;0"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Bill Type is required"
                    DefaultValue ="\"BankBill\""
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =36
                    LayoutCachedTop =48
                    LayoutCachedWidth =1992
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14592
                    Top =48
                    Width =648
                    Height =300
                    TabIndex =13
                    ForeColor =-2147483615
                    Name ="BillID"
                    ControlSource ="BillID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =14592
                    LayoutCachedTop =48
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15960
                    Top =60
                    Width =435
                    Height =225
                    FontSize =8
                    TabIndex =1
                    BorderColor =0
                    ForeColor =4754549
                    Name ="BillSchRef"
                    ControlSource ="BillSchRef"
                    StatusBarText ="link to schedule table"

                    LayoutCachedLeft =15960
                    LayoutCachedTop =60
                    LayoutCachedWidth =16395
                    LayoutCachedHeight =285
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5064
                    Top =48
                    Width =528
                    Height =300
                    TabIndex =6
                    BorderColor =12632256
                    Name ="SchFSL"
                    ControlSource ="SchFSL"
                    StatusBarText ="FSL Lease"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =5064
                    LayoutCachedTop =48
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16395
                    Top =60
                    Width =345
                    Height =225
                    FontSize =8
                    TabIndex =2
                    ForeColor =4754549
                    Name ="GroupId"
                    ControlSource ="GroupId"
                    StatusBarText ="link to schedule table"

                    LayoutCachedLeft =16395
                    LayoutCachedTop =60
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =285
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9768
                    Top =48
                    Width =396
                    Height =300
                    TabIndex =10
                    BorderColor =12632256
                    Name ="DueonRec"
                    ControlSource ="DueonRec"
                    StatusBarText ="bill due on receipt"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =9768
                    LayoutCachedTop =48
                    LayoutCachedWidth =10164
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15600
                    Top =60
                    Width =360
                    Height =225
                    FontSize =8
                    ForeColor =4754549
                    Name ="MLNo"
                    ControlSource ="MLNo"

                    LayoutCachedLeft =15600
                    LayoutCachedTop =60
                    LayoutCachedWidth =15960
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15312
                    Top =48
                    Width =348
                    Height =300
                    FontSize =8
                    TabIndex =14
                    ForeColor =4754549
                    Name ="BillTypeID"
                    ControlSource ="BillTypeID"
                    DefaultValue ="3"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =15312
                    LayoutCachedTop =48
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =16740
                    Top =60
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =15
                    ForeColor =8210719
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff00000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xd5abc3ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff00000000727272ff727272ff727272ff727272ff00000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff
                    End

                    LayoutCachedLeft =16740
                    LayoutCachedTop =60
                    LayoutCachedWidth =17835
                    LayoutCachedHeight =330
                    PictureCaptionArrangement =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderWidth =1
                    BorderColor =8210719
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =8210719
                    PressedForeColor =8210719
                    Overlaps =1
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
' See "NewBill.cls"
