Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17190
    DatasheetFontHeight =11
    ItemSuffix =439
    Left =4920
    Top =7635
    Right =22155
    Bottom =15735
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x5de8e81cf753e640
    End
    RecordSource ="Select * From Location_Contacts Where ContactId=4639"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =0
            BackColor =15709952
            Name ="FormHeader"
        End
        Begin Section
            Height =6390
            Name ="Detail"
            Begin
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =13695
                    Top =435
                    Width =375
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BackColor =15709952
                    Name ="Label84"
                    Caption ="   "
                    LayoutCachedLeft =13695
                    LayoutCachedTop =435
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =705
                    BorderThemeColorIndex =8
                    ForeThemeColorIndex =2
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9864
                    Top =2625
                    Width =285
                    TabIndex =25
                    Name ="contactOptOut"
                    ControlSource ="contactOptOut"
                    StatusBarText ="email opt out"

                    LayoutCachedLeft =9864
                    LayoutCachedTop =2625
                    LayoutCachedWidth =10149
                    LayoutCachedHeight =2865
                    RowStart =8
                    RowEnd =8
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =1425
                    Width =1725
                    FontSize =8
                    TabIndex =16
                    BorderColor =11830108
                    Name ="contactPrimaryPhone"
                    ControlSource ="contactPrimaryPhone"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =1425
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =1665
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =1425
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Pri Phone:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =1425
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =1665
                            RowStart =4
                            RowEnd =4
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7455
                    Top =1425
                    Width =660
                    FontSize =8
                    TabIndex =17
                    BorderColor =11830108
                    Name ="contactPhoneExt"
                    ControlSource ="contactPhoneExt"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7455
                    LayoutCachedTop =1425
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =1665
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6795
                            Top =1425
                            Width =600
                            Height =240
                            FontSize =9
                            Name ="Label31"
                            Caption ="Ext:"
                            LayoutCachedLeft =6795
                            LayoutCachedTop =1425
                            LayoutCachedWidth =7395
                            LayoutCachedHeight =1665
                            RowStart =4
                            RowEnd =4
                            ColumnStart =3
                            ColumnEnd =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =1725
                    Width =1725
                    FontSize =8
                    TabIndex =19
                    BorderColor =11830108
                    Name ="contactPhone2"
                    ControlSource ="contactPhone2"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =1725
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =1965
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =1725
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label37"
                            Caption ="Alt Phone:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =1725
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =1965
                            RowStart =5
                            RowEnd =5
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =2025
                    Width =1725
                    FontSize =8
                    TabIndex =20
                    BorderColor =11830108
                    Name ="contactCell"
                    ControlSource ="contactCell"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =2025
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =2265
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =2025
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label38"
                            Caption ="Cellular:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =2025
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =2265
                            RowStart =6
                            RowEnd =6
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10650
                    Top =1575
                    Width =4740
                    Height =1653
                    FontSize =8
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =11830108
                    Name ="contactNote"
                    ControlSource ="contactNote"
                    FontName ="Tahoma"

                    LayoutCachedLeft =10650
                    LayoutCachedTop =1575
                    LayoutCachedWidth =15390
                    LayoutCachedHeight =3228
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =10651
                            Top =1380
                            Width =480
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label39"
                            Caption ="Notes:"
                            LayoutCachedLeft =10651
                            LayoutCachedTop =1380
                            LayoutCachedWidth =11131
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =5010
                    Top =2925
                    Width =2385
                    FontSize =9
                    TabIndex =26
                    BackColor =13434879
                    BorderColor =967423
                    Name ="contactClassification"
                    ControlSource ="contactClassification"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ContactType.ContactType FROM ContactType;"
                    ColumnWidths ="1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Classification is Required"
                    AfterUpdate ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =2925
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =3165
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =3
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =2925
                            Width =885
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label3"
                            Caption ="Class¹ :"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =2925
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =3165
                            RowStart =9
                            RowEnd =9
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =2325
                    Width =1725
                    FontSize =8
                    TabIndex =22
                    BorderColor =11830108
                    Name ="contactFax"
                    ControlSource ="contactFax"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =2325
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =2565
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =2325
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label48"
                            Caption ="Fax:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =2325
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =2565
                            RowStart =7
                            RowEnd =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4668
                    Top =5820
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =27
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="3840"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4668
                    LayoutCachedTop =5820
                    LayoutCachedWidth =5244
                    LayoutCachedHeight =6015
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =4020
                            Top =5820
                            Width =634
                            Height =195
                            FontSize =8
                            Name ="Label26"
                            Caption ="clientid:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4020
                            LayoutCachedTop =5820
                            LayoutCachedWidth =4654
                            LayoutCachedHeight =6015
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6060
                    Top =5820
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =28
                    Name ="contactID"
                    ControlSource ="contactID"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6060
                    LayoutCachedTop =5820
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =6015
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =5256
                            Top =5820
                            Width =792
                            Height =195
                            FontSize =8
                            Name ="Label43"
                            Caption ="contactID:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =5256
                            LayoutCachedTop =5820
                            LayoutCachedWidth =6048
                            LayoutCachedHeight =6015
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =11
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7455
                    Top =2925
                    Width =2895
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label55"
                    Caption ="¹ Required or Must be Selected"
                    LayoutCachedLeft =7455
                    LayoutCachedTop =2925
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =3165
                    RowStart =9
                    RowEnd =9
                    ColumnStart =4
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =2625
                    Width =3105
                    FontSize =8
                    TabIndex =24
                    BorderColor =11830108
                    Name ="contactEmail"
                    ControlSource ="contactEmail"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =2625
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =2865
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =2625
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label35"
                            Caption ="E-mail:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =2625
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =2865
                            RowStart =8
                            RowEnd =8
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5865
                    Top =495
                    Width =1530
                    Height =270
                    FontSize =8
                    TabIndex =2
                    BorderColor =11830108
                    Name ="contactFirst"
                    ControlSource ="contactFirst"
                    ValidationRule ="Is Not Null"
                    ValidationText ="First Name is required"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5865
                    LayoutCachedTop =495
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =765
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7455
                    Top =495
                    Width =660
                    Height =270
                    FontSize =8
                    TabIndex =3
                    BorderColor =11830108
                    Name ="contactMiddle"
                    ControlSource ="contactMiddle"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7455
                    LayoutCachedTop =495
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8175
                    Top =495
                    Width =1395
                    Height =270
                    FontSize =8
                    TabIndex =4
                    BorderColor =11830108
                    Name ="contactLast"
                    ControlSource ="contactLast"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Last Name is required."
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8175
                    LayoutCachedTop =495
                    LayoutCachedWidth =9570
                    LayoutCachedHeight =765
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =1125
                    Width =3105
                    FontSize =8
                    TabIndex =14
                    BorderColor =11830108
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =1125
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =1125
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label71"
                            Caption ="Title:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =1125
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =1365
                            RowStart =3
                            RowEnd =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8175
                    Top =825
                    Width =1395
                    Height =240
                    FontSize =8
                    TabIndex =11
                    Name ="MakeFullContactBtn"
                    Caption ="◄ update"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8175
                    LayoutCachedTop =825
                    LayoutCachedWidth =9570
                    LayoutCachedHeight =1065
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverThemeColorIndex =2
                    QuickStyle =9
                    QuickStyleMask =-241
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =825
                    Width =3105
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483615
                    BorderColor =15062992
                    ForeColor =16777215
                    Name ="contactFullName"
                    ControlSource ="contactFullName"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5010
                    LayoutCachedTop =825
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =1065
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =825
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label42"
                            Caption ="Full Name:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =825
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =1065
                            RowStart =2
                            RowEnd =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13755
                    Top =975
                    Width =1275
                    FontSize =9
                    TabIndex =12
                    Name ="ConObsoleteDt"
                    ControlSource ="ContactObsoletedt"
                    Format ="Short Date"
                    StatusBarText ="Double Click the Unlock button to unlock this field"
                    ControlTipText ="Double Click the Unlock button to unlock this field"

                    LayoutCachedLeft =13755
                    LayoutCachedTop =975
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =2
                            Left =13695
                            Top =735
                            Width =1425
                            Height =525
                            FontSize =10
                            FontWeight =700
                            BackColor =12493372
                            ForeColor =16777215
                            Name ="Label50"
                            Caption ="As of Date:"
                            LayoutCachedLeft =13695
                            LayoutCachedTop =735
                            LayoutCachedWidth =15120
                            LayoutCachedHeight =1260
                            BorderThemeColorIndex =8
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5010
                    Top =495
                    Width =795
                    Height =270
                    FontSize =8
                    TabIndex =1
                    BorderColor =11830108
                    Name ="contactPrefix"
                    ControlSource ="contactPrefix"
                    RowSourceType ="Value List"
                    RowSource ="Dr.;Mr.;Mrs.;Ms."
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =495
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4065
                            Top =495
                            Width =885
                            Height =270
                            FontSize =9
                            Name ="Label28"
                            Caption ="Name:"
                            LayoutCachedLeft =4065
                            LayoutCachedTop =495
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9630
                    Top =495
                    Width =720
                    Height =270
                    FontSize =8
                    TabIndex =5
                    BorderColor =11830108
                    Name ="contactSuffix"
                    ControlSource ="contactSuffix"
                    RowSourceType ="Value List"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9630
                    LayoutCachedTop =495
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8220
                    Top =1350
                    Width =1296
                    FontSize =9
                    TabIndex =15
                    Name ="LastChangeDt"
                    ControlSource ="LastChangeDt"
                    StatusBarText ="updated when contact unlocked to make change"

                    LayoutCachedLeft =8220
                    LayoutCachedTop =1350
                    LayoutCachedWidth =9516
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =163
                            TextAlign =2
                            Left =8175
                            Top =1110
                            Width =1395
                            Height =555
                            FontSize =9
                            BackColor =11525325
                            Name ="Label69"
                            Caption ="Last Updated"
                            LayoutCachedLeft =8175
                            LayoutCachedTop =1110
                            LayoutCachedWidth =9570
                            LayoutCachedHeight =1665
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =12735
                    Top =525
                    Width =720
                    Height =660
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="Unlockit"
                    Caption ="Unlock"
                    StatusBarText ="Double Click to Unlock"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double Click to Unlock"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12735
                    LayoutCachedTop =525
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =1185
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =8
                    BorderThemeColorIndex =8
                    BorderShade =50.0
                    HoverThemeColorIndex =7
                    QuickStyle =34
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5010
                    Top =240
                    Width =795
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label29"
                    Caption ="Salutation"
                    LayoutCachedLeft =5010
                    LayoutCachedTop =240
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =435
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5865
                    Top =240
                    Width =1530
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label30"
                    Caption ="First¹"
                    LayoutCachedLeft =5865
                    LayoutCachedTop =240
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =435
                    ColumnStart =2
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =7455
                    Top =240
                    Width =660
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label34"
                    Caption ="Middle"
                    LayoutCachedLeft =7455
                    LayoutCachedTop =240
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =435
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =8175
                    Top =240
                    Width =1395
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label32"
                    Caption ="Last¹"
                    LayoutCachedLeft =8175
                    LayoutCachedTop =240
                    LayoutCachedWidth =9570
                    LayoutCachedHeight =435
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =12540
                    Top =330
                    Width =2820
                    Height =1005
                    Name ="Box75"
                    LayoutCachedLeft =12540
                    LayoutCachedTop =330
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =1335
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =570
                    Width =3825
                    Height =5430
                    FontSize =8
                    TabIndex =8
                    Name ="ContactsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT contactID, clientRef, contactClassification As Classification, contactFul"
                        "lName AS Contact FROM Location_contacts WHERE clientRef=4400 Order by contactFir"
                        "st, contactLast ;"
                    ColumnWidths ="0;0;1728;2088"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =570
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =6000
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =30
                    Width =2835
                    Height =465
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label80"
                    Caption ="Contacts"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =495
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10980
                    Top =675
                    Width =1395
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="AddContact"
                    Caption ="Add Contact"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10980
                    LayoutCachedTop =675
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =930
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =10980
                    Top =345
                    Width =1395
                    Height =255
                    FontSize =8
                    FontWeight =700
                    Name ="SaveUpdate"
                    Caption ="Save | Update"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10980
                    LayoutCachedTop =345
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =12780
                    Top =5760
                    Width =1725
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =29
                    Name ="DeleteContact"
                    Caption ="DELETE  This Contact"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12780
                    LayoutCachedTop =5760
                    LayoutCachedWidth =14505
                    LayoutCachedHeight =6000
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2366701
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10980
                    Top =990
                    Width =1395
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="Printcontacts"
                    Caption ="Print Contacts"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10980
                    LayoutCachedTop =990
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =1245
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9630
                    Top =240
                    Width =720
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label36"
                    Caption ="Suffix"
                    LayoutCachedLeft =9630
                    LayoutCachedTop =240
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =435
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8175
                    Top =2625
                    Width =1665
                    Height =240
                    FontSize =9
                    Name ="Label25"
                    Caption ="Opt Out of Email"
                    LayoutCachedLeft =8175
                    LayoutCachedTop =2625
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =2865
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =247
                    Left =13770
                    Top =495
                    TabIndex =6
                    Name ="ContactObsolete"
                    ControlSource ="ContactObsolete"
                    StatusBarText ="Double Click the Unlock button to unlock this field"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double Click the Unlock button to unlock this field"

                    LayoutCachedLeft =13770
                    LayoutCachedTop =495
                    LayoutCachedWidth =14030
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =2
                            Left =14055
                            Top =435
                            Width =1050
                            Height =270
                            FontSize =10
                            BackColor =15709952
                            Name ="Label49"
                            Caption ="   Obsolete "
                            LayoutCachedLeft =14055
                            LayoutCachedTop =435
                            LayoutCachedWidth =15105
                            LayoutCachedHeight =705
                            BorderThemeColorIndex =8
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9864
                    Top =2325
                    Width =285
                    TabIndex =23
                    Name ="contactCofA"
                    ControlSource ="contactCofA"
                    StatusBarText ="email opt out"

                    LayoutCachedLeft =9864
                    LayoutCachedTop =2325
                    LayoutCachedWidth =10149
                    LayoutCachedHeight =2565
                    RowStart =8
                    RowEnd =8
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8175
                    Top =2325
                    Width =1665
                    Height =240
                    FontSize =9
                    Name ="Label398"
                    Caption ="COA Signator Only"
                    LayoutCachedLeft =8175
                    LayoutCachedTop =2325
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =2565
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9864
                    Top =2025
                    Width =285
                    TabIndex =21
                    Name ="contactInc"
                    ControlSource ="contactInc"
                    StatusBarText ="email opt out"

                    LayoutCachedLeft =9864
                    LayoutCachedTop =2025
                    LayoutCachedWidth =10149
                    LayoutCachedHeight =2265
                    RowStart =8
                    RowEnd =8
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8175
                    Top =2025
                    Width =1665
                    Height =240
                    FontSize =9
                    Name ="Label400"
                    Caption ="Incumbent Signator"
                    LayoutCachedLeft =8175
                    LayoutCachedTop =2025
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =2265
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =8
                    ListWidth =12960
                    Left =4020
                    Top =3660
                    Width =6480
                    FontSize =9
                    TabIndex =30
                    BorderColor =967423
                    Name ="LocationCBO"
                    ControlSource ="cLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationID, LocationName, LocationStreet, LAddrLine1, LAddrLine2, LCity, "
                        "LState, LZip From Locations WHERE ClientGroupID=15 AND LocObsolete = False;"
                    ColumnWidths ="0;2160;2160;2160;1800;1800;504;7200"
                    StatusBarText ="lookup to Location Table"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =4020
                    LayoutCachedTop =3660
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =3900
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =4020
                            Top =3420
                            Width =3120
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label422"
                            Caption ="Select Addess"
                            LayoutCachedLeft =4020
                            LayoutCachedTop =3420
                            LayoutCachedWidth =7140
                            LayoutCachedHeight =3660
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10695
                    Top =4275
                    Width =3720
                    Height =1305
                    FontSize =8
                    TabIndex =31
                    BorderColor =0
                    ForeColor =-2147483617
                    Name ="ContactDetail"

                    LayoutCachedLeft =10695
                    LayoutCachedTop =4275
                    LayoutCachedWidth =14415
                    LayoutCachedHeight =5580
                    ThemeFontIndex =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =10620
                    Top =4080
                    Width =3900
                    Height =1560
                    Name ="Box656"
                    LayoutCachedLeft =10620
                    LayoutCachedTop =4080
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =5640
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =10650
                    Top =3960
                    Width =1965
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="Label657"
                    Caption ="Contact - Location Detail"
                    LayoutCachedLeft =10650
                    LayoutCachedTop =3960
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =4200
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =4020
                    Top =4020
                    Width =6480
                    Height =1680
                    FontSize =9
                    TabIndex =32
                    BoundColumn =1
                    Name ="LinkedLocationsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationID, ContactID, LinkType, Locationsn & ' (' & [LocationType] & ')'"
                        " AS Location, City & ', ' & State AS City-State, LinkTypeID FROM ContactsLinkedT"
                        "oLocationQuery WHERE C.ContactID =4639 ORDER BY Locationsn AND LocationType;"
                    ColumnWidths ="0;0;1296;2592;2592;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =4020
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =5700
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2220
                            Top =3900
                            Width =765
                            Height =315
                            Name ="Label436"
                            Caption ="List435:"
                            LayoutCachedLeft =2220
                            LayoutCachedTop =3900
                            LayoutCachedWidth =2985
                            LayoutCachedHeight =4215
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =8040
                    Top =5760
                    Width =2460
                    Height =270
                    FontSize =8
                    TabIndex =33
                    Name ="UnLinkedContact"
                    Caption ="Un-Link Selected Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8040
                    LayoutCachedTop =5760
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =6030
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    UseTheme =1
                    BorderColor =16777215
                    HoverColor =1643706
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    HoverForeShade =65.0
                    PressedForeThemeColorIndex =1
                    PressedForeShade =65.0
                    QuickStyle =22
                    QuickStyleMask =-629
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
CodeBehindForm
' See "ClientContact_sub.cls"
