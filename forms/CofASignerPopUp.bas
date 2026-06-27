Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =62
    Left =5670
    Top =3315
    Right =12870
    Bottom =10440
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf0eb4be7365fe540
    End
    RecordSource ="SELECT TOP 1 vw_SixKeys.ClientGroupID, vw_SixKeys.clientID, vw_SixKeys.SchID, vw"
        "_SixKeys.GroupID, vw_SixKeys.AssetsInGroup, vw_SixKeys.Schedule, vw_SixKeys.Unit"
        "Group FROM vw_SixKeys WHERE (((vw_SixKeys.GroupID)=2779));"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            Height =435
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =45
                    Width =5760
                    Height =435
                    ColumnOrder =0
                    FontSize =12
                    TabIndex =2
                    BorderColor =0
                    ForeColor =10040879
                    Name ="Label6"
                    ControlSource ="=\"CofA Signer: Sch:\" & [Schedule] & \" Grp:\" & [UnitGroup]"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =5775
                    Top =60
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

                    LayoutCachedLeft =5775
                    LayoutCachedTop =60
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =288
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
                    OverlapFlags =247
                    PictureType =1
                    Left =6346
                    Top =60
                    Width =750
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

                    LayoutCachedLeft =6346
                    LayoutCachedTop =60
                    LayoutCachedWidth =7096
                    LayoutCachedHeight =288
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6705
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =4590
                    Top =780
                    Width =255
                    Height =240
                    Name ="Box48"
                    LayoutCachedLeft =4590
                    LayoutCachedTop =780
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =1020
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =4632
                    Top =825
                    TabIndex =6
                    Name ="AskForDate"
                    ControlTipText ="Will prompt you for a Date to use"
                    GridlineColor =0

                    LayoutCachedLeft =4632
                    LayoutCachedTop =825
                    LayoutCachedWidth =4892
                    LayoutCachedHeight =1065
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =3315
                            Top =750
                            Width =1230
                            Height =285
                            FontSize =10
                            Name ="Label90"
                            Caption ="Ask For Date "
                            LayoutCachedLeft =3315
                            LayoutCachedTop =750
                            LayoutCachedWidth =4545
                            LayoutCachedHeight =1035
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2355
                    Top =6300
                    Width =555
                    Height =210
                    FontSize =10
                    TabIndex =9
                    ForeColor =5026082
                    Name ="UnitsInGroup"
                    ControlSource ="AssetsInGroup"

                    LayoutCachedLeft =2355
                    LayoutCachedTop =6300
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =1800
                    Top =6300
                    Width =555
                    Height =210
                    FontSize =10
                    TabIndex =2
                    ForeColor =5026082
                    Name ="GroupID"
                    ControlSource ="GroupID"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =6300
                    LayoutCachedWidth =2355
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =690
                    Top =6300
                    Width =555
                    Height =210
                    FontSize =10
                    TabIndex =1
                    ForeColor =5026082
                    Name ="ClientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =690
                    LayoutCachedTop =6300
                    LayoutCachedWidth =1245
                    LayoutCachedHeight =6510
                End
                Begin Subform
                    OverlapFlags =85
                    Left =105
                    Top =1560
                    Width =7095
                    Height =4485
                    TabIndex =3
                    Name ="Schedule_CoLessee"
                    SourceObject ="Form.Schedule_CoLessee"

                    LayoutCachedLeft =105
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =6045
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =135
                    Top =6300
                    Width =555
                    Height =210
                    FontSize =10
                    TabIndex =4
                    ForeColor =5026082
                    Name ="ClientGroupId"
                    ControlSource ="ClientGroupID"

                    LayoutCachedLeft =135
                    LayoutCachedTop =6300
                    LayoutCachedWidth =690
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1245
                    Top =6300
                    Width =555
                    Height =210
                    FontSize =10
                    TabIndex =5
                    ForeColor =5026082
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =6300
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =6510
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =6912
                    Left =1695
                    Top =345
                    Width =5160
                    Height =270
                    FontSize =10
                    Name ="CofaSignCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.contactID, Contacts.contactFullName AS Name, Contacts.contactTit"
                        "le AS Title, Contacts.contactInc, Contacts.contactCofA FROM Contacts WHERE (((Co"
                        "ntacts.ContactObsolete)<>-1) AND ((Contacts.clientRef)=[Forms]![CofASignerPopUp]"
                        "![clientID]));"
                    ColumnWidths ="0;2880;2880;576;576"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1695
                    LayoutCachedTop =345
                    LayoutCachedWidth =6855
                    LayoutCachedHeight =615
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =345
                            Width =1575
                            Height =285
                            FontSize =10
                            Name ="Label1"
                            Caption ="Set/Change Signer"
                            LayoutCachedLeft =60
                            LayoutCachedTop =345
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =630
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5310
                    Top =780
                    Width =1575
                    FontSize =10
                    TabIndex =7
                    Name ="MakeCofA"
                    Caption ="Run CofA"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run The C of A as set"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5310
                    LayoutCachedTop =780
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =1140
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackColor =9732411
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3780
                    Top =6300
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =8
                    ForeColor =5026082
                    Name ="GroupNo"
                    ControlSource ="UnitGroup"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =6300
                    LayoutCachedWidth =4650
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2910
                    Top =6300
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =10
                    ForeColor =5026082
                    Name ="Text43"
                    ControlSource ="Schedule"

                    LayoutCachedLeft =2910
                    LayoutCachedTop =6300
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Width =6795
                    Height =285
                    FontSize =12
                    TabIndex =11
                    BorderColor =0
                    Name ="Text44"
                    ControlSource ="=\"Current Signer: \" & Nz([SignerName],\"None\")"

                    LayoutCachedLeft =165
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =2430
                    Top =765
                    Width =255
                    Height =240
                    Name ="Box49"
                    LayoutCachedLeft =2430
                    LayoutCachedTop =765
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =1005
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2472
                    Top =810
                    TabIndex =12
                    BorderColor =16777215
                    Name ="HideSignatureDate"
                    ControlTipText ="Hide the word 'Date:' on Page 1"
                    GridlineColor =0

                    LayoutCachedLeft =2472
                    LayoutCachedTop =810
                    LayoutCachedWidth =2732
                    LayoutCachedHeight =1050
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =105
                            Top =750
                            Width =2280
                            Height =285
                            FontSize =10
                            Name ="Label51"
                            Caption ="Hide Signature Date Page 1"
                            LayoutCachedLeft =105
                            LayoutCachedTop =750
                            LayoutCachedWidth =2385
                            LayoutCachedHeight =1035
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =2430
                    Top =1065
                    Width =255
                    Height =240
                    Name ="Box55"
                    LayoutCachedLeft =2430
                    LayoutCachedTop =1065
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =1305
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2472
                    Top =1110
                    TabIndex =13
                    BorderColor =16777215
                    Name ="HideAcceptedDate"
                    ControlTipText ="Hide the 'Accepted Date:' on Page 2"
                    GridlineColor =0

                    LayoutCachedLeft =2472
                    LayoutCachedTop =1110
                    LayoutCachedWidth =2732
                    LayoutCachedHeight =1350
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =105
                            Top =1050
                            Width =2280
                            Height =285
                            FontSize =10
                            Name ="Label57"
                            Caption ="Hide Accepted Date Page 2"
                            LayoutCachedLeft =105
                            LayoutCachedTop =1050
                            LayoutCachedWidth =2385
                            LayoutCachedHeight =1335
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5085
                    Top =6195
                    Width =1905
                    FontSize =10
                    TabIndex =14
                    Name ="MakeCofASchedule"
                    Caption ="Run Schedule  CofA"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run The C of A as set"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5085
                    LayoutCachedTop =6195
                    LayoutCachedWidth =6990
                    LayoutCachedHeight =6555
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =4590
                    Top =1065
                    Width =255
                    Height =240
                    Name ="Box59"
                    LayoutCachedLeft =4590
                    LayoutCachedTop =1065
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =1305
                    BackThemeColorIndex =9
                    BackShade =75.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =4632
                    Top =1110
                    TabIndex =15
                    Name ="DeliveryAsAccepted"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Will prompt you for a Date to use"
                    GridlineColor =0

                    LayoutCachedLeft =4632
                    LayoutCachedTop =1110
                    LayoutCachedWidth =4892
                    LayoutCachedHeight =1350
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextAlign =3
                            Left =2685
                            Top =1035
                            Width =1875
                            Height =285
                            FontSize =10
                            Name ="Label61"
                            Caption =" Delivery as Accepted "
                            LayoutCachedLeft =2685
                            LayoutCachedTop =1035
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =1320
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                        End
                    End
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
' See "CofASignerPopUp.cls"
