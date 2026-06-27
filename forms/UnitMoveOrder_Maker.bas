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
    RecordLocks =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =76
    Left =4845
    Top =3315
    Right =24630
    Bottom =16455
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x1f1099887334e640
    End
    RecordSource ="Select * From UnitMoveOrder Where mAssetID =308054"
    OnClose ="[Event Procedure]"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin FormHeader
            Height =735
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =540
                    Width =18000
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =735
                    TabIndex =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =120
                    Width =8370
                    Height =720
                    FontSize =24
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Unit Transport / Driveaway Manager"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =720
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =16140
                    Top =180
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

                    LayoutCachedLeft =16140
                    LayoutCachedTop =180
                    LayoutCachedWidth =16695
                    LayoutCachedHeight =408
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
                    PictureType =1
                    Left =16711
                    Top =180
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

                    LayoutCachedLeft =16711
                    LayoutCachedTop =180
                    LayoutCachedWidth =17461
                    LayoutCachedHeight =408
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
            Height =9840
            Name ="Detail"
            AlternateBackColor =15527148
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    OverlapFlags =87
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =120
                    Top =3893
                    Width =5472
                    Height =3000
                    FontSize =10
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT mOldLocationID as LocationID, PickupLocation as Pickup, mStatus, Count(As"
                        "setID) as Units FROM vw_UnitMoveDashboard AS M WHERE [Status]='EG' Or [Status] ="
                        "'EM' AND mClientID=5204 Group BY mOldLocationID, PickupLocation, mStatus ORDER B"
                        "Y PickupLocation;"
                    ColumnWidths ="0;2880;1728;720"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =3893
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =6893
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =161
                    Left =17400
                    Top =360
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="line2"
                    Caption ="──┐"
                    FontName ="Arial"
                    LayoutCachedLeft =17400
                    LayoutCachedTop =360
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =4
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =161
                    Left =2220
                    Top =3420
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="Label53"
                    Caption ="┌──"
                    FontName ="Arial"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =3420
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =3675
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =600
                    Width =3825
                    Height =2520
                    FontSize =10
                    TabIndex =1
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.ClientID, vw_SixKeys.ClientGroupID, vw_SixKeys.ClientShortName"
                        " AS Client, Count(vw_SixKeys.AssetID) AS Units FROM vw_SixKeys WHERE (((vw_SixKe"
                        "ys.Status)='EG' Or (vw_SixKeys.Status)='EM')) OR (((vw_SixKeys.Status)='EZ') AND"
                        " ((vw_SixKeys.ParkRepairID) Is Null)) GROUP BY vw_SixKeys.ClientID, vw_SixKeys.C"
                        "lientGroupID, vw_SixKeys.ClientShortName;"
                    ColumnWidths ="0;0;2880;360"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =600
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =75
                    Top =270
                    Width =3825
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="Search_Client"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =270
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =510
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =75
                            Top =60
                            Width =1350
                            Height =225
                            FontSize =9
                            FontWeight =700
                            Name ="Label45"
                            Caption ="Search Client:"
                            LayoutCachedLeft =75
                            LayoutCachedTop =60
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =285
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =5700
                    Top =600
                    Width =12243
                    Height =6293
                    FontSize =10
                    TabIndex =3
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT AssetID, UnitNum, VIN, [Status],[MStatus], UnitDescShortVer as uDesc, mSt"
                        "artDate as Start, DeliveryLocation as [Deliver To] FROM vw_UnitMoveDashboard WHE"
                        "RE ([Status]='EG' Or [Status] ='EM') AND mOldLocationID=1378 ORDER BY mStartDate"
                        ", UnitNum,VIN;"
                    ColumnWidths ="0;1152;2160;432;1296;4320;1296;2592"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5700
                    LayoutCachedTop =600
                    LayoutCachedWidth =17943
                    LayoutCachedHeight =6893
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11970
                    Top =240
                    Width =450
                    Height =255
                    FontSize =9
                    TabIndex =4
                    ForeColor =-2147483615
                    Name ="UnitListcount"
                    ControlSource ="=[UnitsList].ListCount-1"

                    LayoutCachedLeft =11970
                    LayoutCachedTop =240
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =495
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =11040
                            Top =240
                            Width =930
                            Height =255
                            FontSize =9
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label49"
                            Caption ="Units In List"
                            LayoutCachedLeft =11040
                            LayoutCachedTop =240
                            LayoutCachedWidth =11970
                            LayoutCachedHeight =495
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =2667
                    Top =3300
                    Width =2895
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="PickUpByLocation"
                    Caption ="Open Pick Up Form for Location"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2667
                    LayoutCachedTop =3300
                    LayoutCachedWidth =5562
                    LayoutCachedHeight =3570
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =15240
                    Top =240
                    Width =2205
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="MoveByUnit"
                    Caption ="Open Pick Up For Selected Unit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =135
                    TopPadding =105
                    RightPadding =150
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15240
                    LayoutCachedTop =240
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =2
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =5880
                    Top =645
                    Width =3975
                    Height =345
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label55"
                    Caption ="◄ First Select A Client"
                    LayoutCachedLeft =5880
                    LayoutCachedTop =645
                    LayoutCachedWidth =9855
                    LayoutCachedHeight =990
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =5880
                    Top =3690
                    Width =3045
                    Height =360
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label56"
                    Caption ="◄Then Select A Location"
                    LayoutCachedLeft =5880
                    LayoutCachedTop =3690
                    LayoutCachedWidth =8925
                    LayoutCachedHeight =4050
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =1200
                    Top =8280
                    Width =4020
                    Height =300
                    FontWeight =700
                    TabIndex =7
                    Name ="mOldLocationID"
                    ControlSource ="mOldLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                        "ete)=No) AND ((Locations.ClientGroupID)=[Forms]![UnitMoveOrder_Maker]![ClientGro"
                        "upID])) OR (((Locations.LocObsolete)=No) AND ((Locations.ClientGroupID)=771)) OR"
                        "DER BY Locations.LocationName, Locations.LocationType;"
                    ColumnWidths ="0;2160;2160;1440;720;720"
                    StatusBarText ="Link to Contacts table"
                    BottomPadding =150

                    LayoutCachedLeft =1200
                    LayoutCachedTop =8280
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =8580
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =7
                    BackTint =40.0
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =8280
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label5170"
                            Caption ="Move From"
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =8280
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =8580
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =6360
                    Top =8280
                    Width =4020
                    Height =300
                    FontWeight =700
                    TabIndex =8
                    Name ="mNewLocationID"
                    ControlSource ="mNewLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                        "ete)=No) AND ((Locations.ClientGroupID)=[Forms]![UnitMoveOrder_Maker]![ClientGro"
                        "upID])) OR (((Locations.ClientGroupID)=771)) ORDER BY Locations.LocationName, Lo"
                        "cations.LocationType;"
                    ColumnWidths ="0;2160;2160;1440;720;720"
                    StatusBarText ="Link to Contacts table"
                    BottomPadding =150

                    LayoutCachedLeft =6360
                    LayoutCachedTop =8280
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =8580
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5280
                            Top =8280
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label5165"
                            Caption ="Move To"
                            BottomPadding =150
                            LayoutCachedLeft =5280
                            LayoutCachedTop =8280
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =8580
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =13500
                    Top =7920
                    Width =2520
                    Height =300
                    TabIndex =9
                    BackColor =16183539
                    Name ="iVendorID"
                    ControlSource ="mVendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorOb"
                        "solete FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.Vendo"
                        "rSNm;"
                    ColumnWidths ="0;10800;2880"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =13500
                    LayoutCachedTop =7920
                    LayoutCachedWidth =16020
                    LayoutCachedHeight =8220
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12780
                            Top =7920
                            Width =660
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label76"
                            Caption ="Vendor"
                            BottomPadding =150
                            LayoutCachedLeft =12780
                            LayoutCachedTop =7920
                            LayoutCachedWidth =13440
                            LayoutCachedHeight =8220
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =8700
                    Width =4020
                    Height =300
                    TabIndex =10
                    Name ="mContact"
                    ControlSource ="mContact"
                    BottomPadding =150

                    LayoutCachedLeft =1200
                    LayoutCachedTop =8700
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =9000
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =8700
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label375"
                            Caption ="Contact "
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =8700
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =9000
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11520
                    Top =7920
                    Width =1140
                    Height =300
                    TabIndex =11
                    Name ="wStartDate"
                    ControlSource ="mStartDate"
                    Format ="Short Date"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"wStartDate\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">AddRmktExpFor"
                        End
                        Begin
                            Comment ="_AXL:m</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =11520
                    LayoutCachedTop =7920
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =8220
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =10500
                            Top =7920
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label108"
                            Caption ="Start Date"
                            LayoutCachedLeft =10500
                            LayoutCachedTop =7920
                            LayoutCachedWidth =11460
                            LayoutCachedHeight =8220
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =7920
                    Width =3168
                    Height =300
                    TabIndex =12
                    Name ="mOrderNum"
                    ControlSource ="mOrderNum"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =7920
                    LayoutCachedWidth =4368
                    LayoutCachedHeight =8220
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =7920
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Move Order "
                            LayoutCachedLeft =120
                            LayoutCachedTop =7920
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =8220
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11520
                    Top =8280
                    Width =1140
                    Height =300
                    TabIndex =13
                    Name ="mEndDate"
                    ControlSource ="mEndDate"
                    Format ="Short Date"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"mEndDate\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">AddRmktExpForm<"
                        End
                        Begin
                            Comment ="_AXL:/Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =11520
                    LayoutCachedTop =8280
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =8580
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =10500
                            Top =8280
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label285"
                            Caption ="Close Date"
                            LayoutCachedLeft =10500
                            LayoutCachedTop =8280
                            LayoutCachedWidth =11460
                            LayoutCachedHeight =8580
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =9060
                    Width =4020
                    Height =300
                    TabIndex =14
                    Name ="mEmail"
                    ControlSource ="mEmail"
                    BottomPadding =150

                    LayoutCachedLeft =1200
                    LayoutCachedTop =9060
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =9360
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =9060
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label394"
                            Caption ="Email "
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =9060
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =9360
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =9420
                    Width =4020
                    Height =300
                    TabIndex =15
                    Name ="mPhone"
                    ControlSource ="mPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    BottomPadding =150

                    LayoutCachedLeft =1200
                    LayoutCachedTop =9420
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =9720
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =9420
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label396"
                            Caption ="Phone "
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =9420
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =9720
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16680
                    Top =7920
                    Width =1020
                    Height =300
                    TabIndex =16
                    Name ="mMoveCost"
                    ControlSource ="mMoveCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    BottomPadding =150

                    LayoutCachedLeft =16680
                    LayoutCachedTop =7920
                    LayoutCachedWidth =17700
                    LayoutCachedHeight =8220
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =16140
                            Top =7920
                            Width =480
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label400"
                            Caption ="Cost "
                            BottomPadding =150
                            LayoutCachedLeft =16140
                            LayoutCachedTop =7920
                            LayoutCachedWidth =16620
                            LayoutCachedHeight =8220
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
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
                    Left =6360
                    Top =7905
                    Width =4020
                    Height =315
                    TabIndex =17
                    Name ="mStatus"
                    ControlSource ="mStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"Not Ordred\";\"Pending\";\"In Transit\";\"Done\""
                    BottomPadding =150

                    LayoutCachedLeft =6360
                    LayoutCachedTop =7905
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =8220
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5280
                            Top =7920
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label398"
                            Caption ="Status "
                            BottomPadding =150
                            LayoutCachedLeft =5280
                            LayoutCachedTop =7920
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =8220
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =7560
                    Width =18000
                    Height =210
                    BorderColor =16777215
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =7560
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =7770
                    TabIndex =30
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =180
                    Top =7200
                    Width =1785
                    Height =495
                    FontSize =16
                    Name ="Label1053"
                    Caption ="Move Order"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =180
                    LayoutCachedTop =7200
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =7695
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =660
                    Width =840
                    FontSize =9
                    TabIndex =18
                    Name ="ClientGroupID"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =660
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =900
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =6945
                    Width =1728
                    Height =240
                    FontSize =9
                    TabIndex =19
                    Name ="Location_AllOrdred"
                    Caption ="Mark All as Ordered"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3000
                    LayoutCachedTop =6945
                    LayoutCachedWidth =4728
                    LayoutCachedHeight =7185
                    ColumnStart =3
                    ColumnEnd =3
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
                    OverlapFlags =85
                    Left =15420
                    Top =8820
                    Width =2160
                    FontSize =9
                    TabIndex =20
                    Name ="SubmitMoveOrder"
                    Caption ="Update Unit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15420
                    LayoutCachedTop =8820
                    LayoutCachedWidth =17580
                    LayoutCachedHeight =9180
                    ColumnStart =3
                    ColumnEnd =3
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
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =9420
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    Name ="mFaUserID"
                    ControlSource ="mFaUserID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =12360
                    LayoutCachedTop =9420
                    LayoutCachedWidth =13140
                    LayoutCachedHeight =9660
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13140
                    Top =9420
                    Width =1020
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
                    Name ="UnitMoveID"
                    ControlSource ="UnitMoveID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =13140
                    LayoutCachedTop =9420
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =9660
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11460
                    Top =9420
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    Name ="mAssetId"
                    ControlSource ="mAssetId"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =11460
                    LayoutCachedTop =9420
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =9660
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6300
                    Top =8760
                    Width =4860
                    Height =960
                    TabIndex =24
                    Name ="mNote"
                    ControlSource ="mNote"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =8760
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =9720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5280
                            Top =8760
                            Width =960
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label152"
                            Caption ="Notes "
                            LayoutCachedLeft =5280
                            LayoutCachedTop =8760
                            LayoutCachedWidth =6240
                            LayoutCachedHeight =9060
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13500
                    Top =8280
                    Width =2520
                    Height =300
                    TabIndex =25
                    Name ="Text69"
                    ControlSource ="moveETA"
                    BottomPadding =150

                    LayoutCachedLeft =13500
                    LayoutCachedTop =8280
                    LayoutCachedWidth =16020
                    LayoutCachedHeight =8580
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12780
                            Top =8280
                            Width =660
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label70"
                            Caption ="ETA"
                            BottomPadding =150
                            LayoutCachedLeft =12780
                            LayoutCachedTop =8280
                            LayoutCachedWidth =13440
                            LayoutCachedHeight =8580
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    DecimalPlaces =0
                    ScrollBars =2
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16680
                    Top =8340
                    Width =1020
                    Height =300
                    TabIndex =26
                    Name ="Text71"
                    ControlSource ="mMiles"
                    Format ="Standard"
                    BottomPadding =150

                    LayoutCachedLeft =16680
                    LayoutCachedTop =8340
                    LayoutCachedWidth =17700
                    LayoutCachedHeight =8640
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =16140
                            Top =8340
                            Width =570
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label72"
                            Caption ="Miles "
                            BottomPadding =150
                            LayoutCachedLeft =16140
                            LayoutCachedTop =8340
                            LayoutCachedWidth =16710
                            LayoutCachedHeight =8625
                            ColumnStart =6
                            ColumnEnd =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =7
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =223
                    Left =120
                    Top =3660
                    Width =2880
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =27
                    ForeColor =8210719
                    Name ="ToggleSortPickup"
                    DefaultValue ="0"
                    Caption ="PickUp Location"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =120
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =3900
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =223
                    Left =3000
                    Top =3660
                    Width =1728
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =28
                    ForeColor =8210719
                    Name ="ToggleSortmStatus"
                    DefaultValue ="0"
                    Caption ="Order Status"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =3660
                    LayoutCachedWidth =4728
                    LayoutCachedHeight =3900
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =223
                    Left =4728
                    Top =3660
                    Width =864
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =29
                    ForeColor =8210719
                    Name ="ToggleSortUnits"
                    DefaultValue ="0"
                    Caption ="Units"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =4728
                    LayoutCachedTop =3660
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =3900
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
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
' See "UnitMoveOrder_Maker.cls"
