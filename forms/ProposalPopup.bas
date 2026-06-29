Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11844
    DatasheetFontHeight =11
    ItemSuffix =84
    Left =25185
    Top =3645
    Right =-26431
    Bottom =12825
    Filter ="ProposalID=48"
    RecSrcDt = Begin
        0x14df14bf55c4e440
    End
    RecordSource ="SELECT BuyersProposal.ProposalID, BuyersProposal.BuyerID, Clients.ClientShNm AS "
        "Buyer, Clients.clientType, BuyersProposal.FirstContactDate, BuyersProposal.First"
        "ContactID, BuyersProposal.SalesPerson, BuyersProposal.ProposalName, BuyersPropos"
        "al.ProposalStatus, BuyersProposal.ProposalNotes, BuyersProposal.LastContactID, B"
        "uyersProposal.LastContactDate, BuyersProposal.SalesPersonID FROM BuyersProposal "
        "INNER JOIN Clients ON BuyersProposal.BuyerID = Clients.clientID;"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Tab
            TextFontFamily =0
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =759
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =360
                    Width =2820
                    Height =495
                    FontSize =18
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Proposal"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =495
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =5
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =360
                    Width =6720
                    Height =375
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    BorderColor =12566463
                    ForeColor =0
                    Name ="BuyerCbo"
                    ControlSource ="=DLookUp(\"ClientShNm\",\"Buyers\",\"ClientID=\" & [buyerID])"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =4

                    LayoutCachedLeft =1020
                    LayoutCachedTop =360
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =735
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =180
                            Top =360
                            Width =780
                            Height =375
                            FontSize =9
                            BorderColor =0
                            Name ="Caller_Label"
                            Caption ="Buyer"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =4
                            LayoutCachedLeft =180
                            LayoutCachedTop =360
                            LayoutCachedWidth =960
                            LayoutCachedHeight =735
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10200
                    Top =360
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="SaveUpdate"
                    Caption ="  Save | Update"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000090a46869 ,
                        0x90a4686000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000090a4684e90a468fc ,
                        0x90a468f690a4682d000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000090a4683690a468f690a468ff ,
                        0x90a468ff90a468db90a4680f0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000090a4682d90a468ea90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468b40000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a4681b90a468e490a468ff90a468ff90a468ae ,
                        0x90a468ba90a468ff90a468ff90a4687e00000000000000000000000000000000 ,
                        0x000000000000000090a4680f90a468d290a468ff90a468ff90a468bd90a46806 ,
                        0x90a4681590a468e790a468ff90a468fc90a46848000000000000000000000000 ,
                        0x000000000000000090a4684290a468ff90a468ff90a468c990a4680f00000000 ,
                        0x0000000090a4683c90a468fc90a468ff90a468ed90a468210000000000000000 ,
                        0x00000000000000000000000090a4685790a468cc90a468150000000000000000 ,
                        0x000000000000000090a4687b90a468ff90a468ff90a468cc90a4680600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a468b790a468ff90a468ff90a4689c00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a4681590a468e190a468ff90a468ff90a46863 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000090a4683c90a468f990a468ff90a468f9 ,
                        0x90a4683300000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000090a4687590a468ff90a468ff ,
                        0x90a4689300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000090a468ab90a46887 ,
                        0x90a4680300000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10200
                    LayoutCachedTop =360
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =615
                    PictureCaptionArrangement =5
                    Alignment =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =10319446
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =8872517
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-117
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7920
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMEMode =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1512
                    Top =1152
                    Width =3600
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =12566463
                    ForeColor =0
                    Name ="EventDueDate"
                    ControlSource ="FirstContactDate"
                    Format ="Long Date"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1

                    LayoutCachedLeft =1512
                    LayoutCachedTop =1152
                    LayoutCachedWidth =5112
                    LayoutCachedHeight =1452
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1152
                            Width =1320
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Due Date_Label"
                            Caption =" Date"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            EventProcPrefix ="Due_Date_Label"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =1152
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =1452
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1512
                    Top =780
                    Width =3600
                    Height =300
                    FontSize =10
                    TabIndex =3
                    BoundColumn =2
                    BorderColor =12566463
                    ForeColor =0
                    Name ="ContactCBO"
                    ControlSource ="FirstContactID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.clientRef, [contactFirst] & \" \" & [contactLast] AS Contact, Co"
                        "ntacts.contactID FROM Contacts WHERE (((Contacts.clientRef)=[Forms]![ProposalPop"
                        "Up]![BuyerID]));"
                    ColumnWidths ="0;4320;0"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1
                    ListItemsEditForm ="Customer Details"
                    InheritValueList =0

                    LayoutCachedLeft =1512
                    LayoutCachedTop =780
                    LayoutCachedWidth =5112
                    LayoutCachedHeight =1080
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
                            Top =780
                            Width =1320
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Label100"
                            Caption ="First Contact"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =780
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =1080
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =1500
                    Width =11160
                    Height =2400
                    FontSize =9
                    TabIndex =7
                    BackColor =15728639
                    BorderColor =967423
                    ForeColor =0
                    Name ="ProposalNotes"
                    ControlSource ="ProposalNotes"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GridlineColor =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =3900
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
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6384
                    Top =780
                    Width =3600
                    Height =300
                    FontSize =10
                    TabIndex =4
                    BoundColumn =2
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Combo0"
                    ControlSource ="LastContactID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.clientRef, [contactFirst] & \" \" & [contactLast] AS Contact, Co"
                        "ntacts.contactID FROM Contacts WHERE (((Contacts.clientRef)=[Forms]![ProposalPop"
                        "Up]![BuyerID]));"
                    ColumnWidths ="0;4320;0"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1
                    ListItemsEditForm ="Customer Details"
                    InheritValueList =0

                    LayoutCachedLeft =6384
                    LayoutCachedTop =780
                    LayoutCachedWidth =9984
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
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
                            Left =5172
                            Top =780
                            Width =1140
                            Height =300
                            FontSize =9
                            BackColor =14277081
                            BorderColor =0
                            Name ="Label1"
                            Caption ="Last Contact"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =1
                            LayoutCachedLeft =5172
                            LayoutCachedTop =780
                            LayoutCachedWidth =6312
                            LayoutCachedHeight =1080
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackShade =85.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMEMode =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6384
                    Top =1152
                    Width =3600
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text20"
                    ControlSource ="LastContactDate"
                    Format ="Long Date"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1

                    LayoutCachedLeft =6384
                    LayoutCachedTop =1152
                    LayoutCachedWidth =9984
                    LayoutCachedHeight =1452
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
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
                            OverlapFlags =85
                            TextAlign =3
                            Left =5172
                            Top =1152
                            Width =1140
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Label21"
                            Caption =" Date"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =1
                            LayoutCachedLeft =5172
                            LayoutCachedTop =1152
                            LayoutCachedWidth =6312
                            LayoutCachedHeight =1452
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1512
                    Top =36
                    Width =8472
                    Height =300
                    FontSize =9
                    BorderColor =12566463
                    ForeColor =0
                    Name ="ProposalName"
                    ControlSource ="ProposalName"
                    StatusBarText ="Double-click to go to details"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1

                    LayoutCachedLeft =1512
                    LayoutCachedTop =36
                    LayoutCachedWidth =9984
                    LayoutCachedHeight =336
                    ColumnStart =1
                    ColumnEnd =3
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
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =36
                            Width =1320
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Label25"
                            Caption ="Proposal Name"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =36
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =336
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1512
                    Top =408
                    Width =3600
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =2
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SalesPerson"
                    ControlSource ="SalesPerson"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TBL_Employees.[User ID], TBL_Employees.[Employee Name] AS [Names], TBL_Em"
                        "ployees.AccessGroup, TBL_Employees.[Active User] FROM TBL_Employees WHERE (((TBL"
                        "_Employees.AccessGroup) Like '*RMKT*' Or (TBL_Employees.AccessGroup) Like '*ITT*"
                        "') AND ((TBL_Employees.[Active User])=True)) ORDER BY TBL_Employees.[Employee Na"
                        "me];"
                    ColumnWidths ="0;1440"
                    StatusBarText ="Double-click to go to details"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =1
                    SeparatorCharacters =255
                    ListItemsEditForm ="Employee Details"

                    LayoutCachedLeft =1512
                    LayoutCachedTop =408
                    LayoutCachedWidth =5112
                    LayoutCachedHeight =708
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
                            Top =408
                            Width =1320
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Assigned To_Label"
                            Caption ="Assigned To"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            EventProcPrefix ="Assigned_To_Label"
                            GroupTable =1
                            LayoutCachedLeft =120
                            LayoutCachedTop =408
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =708
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =300
                    Top =4080
                    Width =5835
                    Height =3720
                    TabIndex =8
                    Name ="EventHistory_Proposal"
                    SourceObject ="Form.EventHistory_Proposal"

                    LayoutCachedLeft =300
                    LayoutCachedTop =4080
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =7800
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =6240
                    Top =4080
                    Width =5250
                    Height =3750
                    FontSize =8
                    TabIndex =9
                    Name ="ProposalUnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM Buyer2Proposal_Query WHERE ProposalID = 5 ;"
                    ColumnWidths ="648;3168;432;0;432;0"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =4080
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =7830
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10545
                    Top =840
                    Width =795
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="LockIt"
                    Caption ="Unlocked"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10545
                    LayoutCachedTop =840
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1080
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10740
                    Top =360
                    Width =420
                    Height =420
                    FontSize =9
                    TabIndex =10
                    ForeColor =0
                    Name ="LockUnLockBtn"
                    Caption ="L"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Click To Un-Lock"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000082c2ea9f82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2ea8100000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff0000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff0000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2ea7b0000000082c2ea7b82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2ea0c0000000082c2ea0c82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2ea7b0000000082c2ea7b82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff00000000 ,
                        0x0000000000000000000000000000000082c2eac682c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eac600000000 ,
                        0x000000000000000000000000000000000000000082c2ea9682c2eaff82c2ea03 ,
                        0x00000000000000000000000082c2ea0382c2eaff82c2ea690000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea8182c2eaff82c2ea24 ,
                        0x00000000000000000000000082c2ea2182c2eaff82c2ea630000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea4882c2eaff82c2ea6c ,
                        0x00000000000000000000000082c2ea6c82c2eaff82c2ea2d0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000082c2eab782c2eaf6 ,
                        0x82c2ea9f82c2ea6682c2ea9f82c2eaf682c2ea9f000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000082c2ea0982c2ea87 ,
                        0x82c2eae482c2eaff82c2eae482c2ea7882c2ea06000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10740
                    LayoutCachedTop =360
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =780
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =52479
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =15527148
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6384
                    Top =408
                    Width =3600
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    BorderColor =12566463
                    ForeColor =0
                    Name ="ProposalStatus"
                    ControlSource ="=DLookUp(\"ProposalStatus\",\"BuyerProposalAutoStatus\",\"ProposalID=\" & [propo"
                        "salID])"
                    DefaultValue ="\"Sent to Buyer\""
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =6384
                    LayoutCachedTop =408
                    LayoutCachedWidth =9984
                    LayoutCachedHeight =708
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
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
                            OverlapFlags =85
                            TextAlign =3
                            Left =5172
                            Top =408
                            Width =1140
                            Height =300
                            FontSize =9
                            BorderColor =0
                            Name ="Status_Label"
                            Caption ="Status"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =1
                            LayoutCachedLeft =5172
                            LayoutCachedTop =408
                            LayoutCachedWidth =6312
                            LayoutCachedHeight =708
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =525
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10092
                    Top =60
                    Width =504
                    FontSize =8
                    BorderColor =12566463
                    Name ="BuyerID"
                    ControlSource ="BuyerID"
                    StatusBarText ="Double-click to go to details"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =9

                    LayoutCachedLeft =10092
                    LayoutCachedTop =60
                    LayoutCachedWidth =10596
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9360
                            Top =60
                            Width =675
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label63"
                            Caption ="BuyerID"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =9
                            LayoutCachedLeft =9360
                            LayoutCachedTop =60
                            LayoutCachedWidth =10035
                            LayoutCachedHeight =300
                            LayoutGroup =3
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11304
                    Top =60
                    Width =504
                    FontSize =8
                    TabIndex =1
                    BorderColor =12566463
                    Name ="ProposalID"
                    ControlSource ="ProposalID"
                    StatusBarText ="Double-click to go to details"
                    FontName ="Segoe UI"
                    Tag ="EditYN"
                    GroupTable =9

                    LayoutCachedLeft =11304
                    LayoutCachedTop =60
                    LayoutCachedWidth =11808
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10656
                            Top =60
                            Width =588
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label65"
                            Caption ="ProposalID"
                            FontName ="Segoe UI"
                            Tag =";RegenerateCaption;"
                            GroupTable =9
                            LayoutCachedLeft =10656
                            LayoutCachedTop =60
                            LayoutCachedWidth =11244
                            LayoutCachedHeight =300
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =9
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ProposalPopup.cls"
