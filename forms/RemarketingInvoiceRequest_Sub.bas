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
    Width =20668
    DatasheetFontHeight =11
    ItemSuffix =94
    Left =-8925
    Top =3795
    Right =12660
    Bottom =14175
    RecSrcDt = Begin
        0x7fc71857537ee540
    End
    RecordSource ="SELECT * FROM RmktInvoiceRequestTbl WHERE (((RmktInvoiceRequestTbl.InvoiceReques"
        "tID)=0));"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =9660
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =0
                    Left =6840
                    Top =5011
                    Width =5805
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    Name ="ClearLabel"
                    Caption ="  ┌─────────┬───────── Click Here to Clear Search Reset List   ─────────┐"
                    FontName ="0S Sans Serif"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0
                    LayoutCachedLeft =6840
                    LayoutCachedTop =5011
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =5281
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =4860
                    Top =5460
                    Width =13860
                    Height =4080
                    FontSize =9
                    TabIndex =31
                    Name ="UnAssignedUnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.Targe"
                        "tSale, RAU.SourceCompany, RAU.[Sch|Grp], RAU.DaysLeft, Count(RAU.InvoiceRequestI"
                        "D) AS Orders FROM (SELECT aU.AssetID, aU.Unitnum, aU.VIN, aU.Description, aU.Sta"
                        "tus, aU.TargetSale, aU.SourceCompany, aU.[Sch|Grp], aU.InvoiceRequestID, DaysLef"
                        "t FROM vw_RMKT_AvailbleUnits AS aU WHERE (((aU.[InvoiceRequestID]) Is Null) AND "
                        "((aU.[AssetID]) Like '**')) OR (((aU.[InvoiceRequestID])<>886) AND ((aU.[AssetID"
                        "]) Like '**')) ORDER BY aU.AssetID)  AS RAU GROUP BY RAU.AssetID, RAU.Unitnum, R"
                        "AU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.SourceCompany, RAU.[Sch"
                        "|Grp], RAU.DaysLeft;"
                    ColumnWidths ="864;864;2160;4320;648;1080;1440;1440;360"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =5460
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =9540
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =60
                    Top =2280
                    Width =4500
                    Height =6840
                    FontSize =9
                    TabIndex =22
                    Name ="InvoiceRequestID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RI.InvoiceRequestID AS Num, RI.InvoiceRequestDate AS Requested, Count(B.U"
                        "nitID) AS Units, RI.BuyerID, RI.InvoiceSentDate AS Sent, RI.InvoicePaidDate AS P"
                        "aid FROM RmktInvoiceRequestTbl AS RI LEFT JOIN RMKT_BuyersOrders AS B ON RI.Invo"
                        "iceRequestID = B.BuyersOrderID WHERE BuyerId= 4263 GROUP BY RI.InvoiceRequestID,"
                        " RI.InvoiceRequestDate, RI.BuyerID, RI.InvoiceSentDate, RI.InvoicePaidDate Order"
                        " by InvoiceRequestID Desc;"
                    ColumnWidths ="576;1008;576;0;1008;1008"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =9120
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =2040
                            Width =2700
                            Height =240
                            FontWeight =700
                            ForeColor =6108695
                            Name ="Label5"
                            Caption ="Select Buyers Order"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =2280
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =120
                    Width =5640
                    Height =315
                    FontWeight =700
                    TabIndex =4
                    ForeColor =6108695
                    Name ="Label8"
                    ControlSource ="=\"Units in Buyers Order Num: \" & [InvoiceRequestID]"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =120
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =435
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =4860
                    Top =4920
                    Width =2760
                    Height =300
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Text9"
                    Caption ="Available Inventory "
                    LayoutCachedLeft =4860
                    LayoutCachedTop =4920
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =5220
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontFamily =0
                    Left =6420
                    Top =4440
                    Width =2574
                    Height =480
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
                    ForeColor =855309
                    Name ="AddToProposal"
                    Caption =" Add Selected Units\015\012 To Buyers Order"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="This will add the selected Units to the Invoice Request"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000072727287727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272870000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272a5727272ff ,
                        0x727272ff727272ff727272ff727272ff727272a5000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727206727272bd ,
                        0x727272ff727272ff727272ff727272bd72727206000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000072727209 ,
                        0x727272c9727272ff727272c97272720900000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x72727215727272b7727272150000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6420
                    LayoutCachedTop =4440
                    LayoutCachedWidth =8994
                    LayoutCachedHeight =4920
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeTint =95.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =2
                    HoverTint =20.0
                    PressedThemeColorIndex =2
                    PressedTint =20.0
                    PressedShade =100.0
                    HoverForeColor =855309
                    HoverForeTint =95.0
                    PressedForeColor =855309
                    PressedForeTint =95.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =100
                    ListWidth =2304
                    Left =2955
                    Top =540
                    Height =237
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    Name ="SalesRep"
                    ControlSource ="SP"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.Rmkt)=Yes) AND ((Emps.OBS"
                        "OLETE)=No)) OR (((Emps.EmpInt)=\"hs\")) ORDER BY Emps.EmpInt;"
                    ColumnWidths ="432;1872"
                    Format =">"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =540
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =777
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1950
                            Top =540
                            Width =999
                            Height =237
                            FontSize =8
                            FontWeight =700
                            Name ="Label18"
                            Caption ="Sales Rep"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =540
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =777
                            LayoutGroup =1
                            BorderThemeColorIndex =2
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =1020
                    Width =1575
                    Height =420
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    ForeColor =0
                    Name ="SaveUpdateOrder"
                    Caption ="  Save | Update\015\012  Buyers Order\015\012"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create an Invoice Request"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ab598980ab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989800000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffd9b4caffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd9b4caffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =1440
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =14211288
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =60
                    Top =9180
                    Width =2274
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =33
                    Name ="DeleteBuyersOrder"
                    Caption ="   Delete  Buyers Order "
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="This will remove the selected Request"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000696969126969698a69696951000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000069696906 ,
                        0x696969a5696969276969696f696969ff696969fc6969694e0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000069696906696969b7 ,
                        0x696969b40000000069696921696969f9696969ff696969f66969693000000000 ,
                        0x000000000000000000000000000000000000000069696906696969b4696969f9 ,
                        0x6969692400000000000000006969697e696969ff696969ff696969e46969691b ,
                        0x0000000000000000000000000000000069696906696969ae696969ff69696984 ,
                        0x00000000000000000000000000000000696969a5696969ff696969ff696969d2 ,
                        0x6969690f000000000000000069696906696969ae696969ff696969d269696906 ,
                        0x0000000000000000000000000000000069696906696969b4696969ff696969ff ,
                        0x696969c66969691269696906696969ab696969ff696969f96969693000000000 ,
                        0x00000000000000000000000000000000000000006969690369696993696969ff ,
                        0x696969ff696969d8696969c0696969ff696969ff696969630000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969695d ,
                        0x696969f3696969ff696969ff696969ff6969699c000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969694e ,
                        0x696969f3696969ff696969ff696969ff69696951000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006969690c69696999696969ff ,
                        0x696969ff696969ff696969cc696969e1696969ff6969698a6969690600000000 ,
                        0x0000000000000000000000000000000069696948696969e1696969ff696969ff ,
                        0x696969ff696969ab696969096969690969696984696969f6696969cc69696921 ,
                        0x0000000000000000000000006969698d696969ff696969ff696969ff696969ff ,
                        0x69696987000000000000000000000000000000006969691e69696993696969ea ,
                        0x696969786969690669696903696969f6696969ff696969ff696969d569696942 ,
                        0x0000000000000000000000000000000000000000000000000000000069696912 ,
                        0x6969696369696927000000006969694e696969b76969696f6969690600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =9180
                    LayoutCachedWidth =2334
                    LayoutCachedHeight =9480
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =9
                    BackTint =100.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    PressedThemeColorIndex =9
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    Width =6180
                    Height =540
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label10"
                    Caption ="Buyers Order Management"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2955
                    Top =780
                    Height =237
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    Name ="RequestDate"
                    ControlSource ="InvoiceRequestDate"
                    Format ="Short Date"
                    ControlTipText ="Set the date the the Order is to be Accepted By"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =780
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1017
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1950
                            Top =780
                            Width =999
                            Height =237
                            FontSize =8
                            FontWeight =700
                            Name ="Label657"
                            Caption ="Request Date"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =780
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =1017
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            BorderThemeColorIndex =2
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4860
                    Top =5220
                    Width =864
                    FontSize =9
                    TabIndex =27
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="AssetSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4860
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =5460
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6588
                    Top =5220
                    Width =2160
                    FontSize =9
                    TabIndex =29
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="VINSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6588
                    LayoutCachedTop =5220
                    LayoutCachedWidth =8748
                    LayoutCachedHeight =5460
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
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8748
                    Top =5220
                    Width =4320
                    FontSize =9
                    TabIndex =30
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="DescSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8748
                    LayoutCachedTop =5220
                    LayoutCachedWidth =13068
                    LayoutCachedHeight =5460
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5724
                    Top =5220
                    Width =864
                    FontSize =9
                    TabIndex =28
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="UnitNumSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5724
                    LayoutCachedTop =5220
                    LayoutCachedWidth =6588
                    LayoutCachedHeight =5460
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =0
                    Left =13080
                    Top =5220
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =12349952
                    Name ="ClearUnAssignedUnits"
                    Caption ="Clear▼"
                    FontName ="0S Sans Serif"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    GridlineColor =0
                    LayoutCachedLeft =13080
                    LayoutCachedTop =5220
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =5445
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =15420
                    Top =4620
                    Width =2574
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =25
                    Name ="PrintRequestedInvoice"
                    Caption ="  Invoice  Request \015\012  INTERNAL ONLY"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Opens the Selected Request's Invoice Request Report"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x00000000000000007272726f727272ea727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff727272ff ,
                        0x727272ea7272726c727272fc727272ff727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff00000000 ,
                        0x727272ff727272f6727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272de7272726672727218000000007272721872727266 ,
                        0x727272de727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ba7272720c90a4686c90a468c390a468ed90a468c390a4686c ,
                        0x7272720c727272ba727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272de787c700f90a468bd90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468bd787c700f727272fc727272ff727272ff000000000000000000000000 ,
                        0x0000000090a4687590a468ff93a76df8e2e7d8ff9daf7aff90a468ff90a468ff ,
                        0x90a468ff90a4687572727284727272fc727272ff00000000727272ff727272ff ,
                        0x7272721b94a86ee792a66bfee2e7d8ffeef1e8ffe7ecdfff93a66cff90a468ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff0691a569fbb5c29affdee4d3ff93a66cffdde3d1ffd1dac1ff90a468ff ,
                        0x90a468ff90a468f900000000000000000000000000000000727272f0ffffffff ,
                        0xffffff1b93a76de590a468ff90a468ff90a468ff98aa73fff2f4edffbac6a1ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff6c90a4688490a468ff90a468ff90a468ff90a468ffa8b788ffe7ecdfff ,
                        0x90a468ff90a4688400000000000000000000000000000000727272ff727272ff ,
                        0x727272e17e866e1790a468cc90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x93a66cd190a46809000000000000000000000000000000000000000000000000 ,
                        0x000000000000000090a4680690a4688190a468d590a468ff90a468d590a46881 ,
                        0x90a4680600000000
                    End

                    LayoutCachedLeft =15420
                    LayoutCachedTop =4620
                    LayoutCachedWidth =17994
                    LayoutCachedHeight =4920
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =6
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedTint =60.0
                    PressedShade =100.0
                    HoverForeThemeColorIndex =3
                    HoverForeTint =100.0
                    HoverForeShade =10.0
                    PressedForeThemeColorIndex =3
                    PressedForeTint =100.0
                    PressedForeShade =10.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    PictureType =1
                    Left =18900
                    Top =5940
                    Width =1674
                    Height =420
                    FontSize =8
                    FontWeight =700
                    TabIndex =32
                    Name ="SearchRequestedForm"
                    Caption ="Search Invoiced Units"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Locate a unit that was previously invoiced (opens another form)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =18900
                    LayoutCachedTop =5940
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =6360
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontFamily =0
                    PictureType =1
                    Left =12420
                    Top =60
                    Width =1674
                    Height =420
                    FontSize =8
                    FontWeight =700
                    Name ="PrevPurchases"
                    Caption ="Prev Purchases"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Open a Report for previous purchased unit by this buyer."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =12420
                    LayoutCachedTop =60
                    LayoutCachedWidth =14094
                    LayoutCachedHeight =480
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =18900
                    Top =5580
                    Width =1674
                    Height =300
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Text20"
                    Caption ="Tools"
                    LayoutCachedLeft =18900
                    LayoutCachedTop =5580
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =5880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13695
                    Top =4860
                    Width =480
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =26
                    Name ="Text47"
                    ControlSource ="=IIf([UnAssignedUnitsList].[ListCount]=0,0,[UnAssignedUnitsList].[ListCount]-1)"
                    FontName ="0S Sans Serif"

                    LayoutCachedLeft =13695
                    LayoutCachedTop =4860
                    LayoutCachedWidth =14175
                    LayoutCachedHeight =5070
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =0
                    Left =13200
                    Top =4860
                    Width =480
                    Height =210
                    FontSize =8
                    FontWeight =700
                    Name ="Label49"
                    Caption ="Units:"
                    FontName ="0S Sans Serif"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =4860
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =5070
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2955
                    Top =1020
                    Height =237
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    Name ="FundingType"
                    ControlSource ="FundingType"
                    RowSourceType ="Value List"
                    RowSource ="\"Cash\";\"Credit-FA\";\"Credit-Outside"
                    DefaultValue ="\"Cash\""
                    ControlTipText ="Payment source for the Buyers Order"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1257
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1950
                            Top =1020
                            Width =999
                            Height =237
                            FontSize =8
                            FontWeight =700
                            Name ="Label22"
                            Caption ="Funding"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =1257
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            BorderThemeColorIndex =2
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =18960
                    Top =720
                    Width =1614
                    Height =480
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="PrintBuyersOrder"
                    Caption ="  BUYERS ORDER"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Opens the Selected Request's Invoice Request Report"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x00000000000000007272726f727272ea727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff727272ff ,
                        0x727272ea7272726c727272fc727272ff727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff00000000 ,
                        0x727272ff727272f6727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272de7272726672727218000000007272721872727266 ,
                        0x727272de727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ba7272720c90a4686c90a468c390a468ed90a468c390a4686c ,
                        0x7272720c727272ba727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272de787c700f90a468bd90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468bd787c700f727272fc727272ff727272ff000000000000000000000000 ,
                        0x0000000090a4687590a468ff93a76df8e2e7d8ff9daf7aff90a468ff90a468ff ,
                        0x90a468ff90a4687572727284727272fc727272ff00000000727272ff727272ff ,
                        0x7272721b94a86ee792a66bfee2e7d8ffeef1e8ffe7ecdfff93a66cff90a468ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff0691a569fbb5c29affdee4d3ff93a66cffdde3d1ffd1dac1ff90a468ff ,
                        0x90a468ff90a468f900000000000000000000000000000000727272f0ffffffff ,
                        0xffffff1b93a76de590a468ff90a468ff90a468ff98aa73fff2f4edffbac6a1ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff6c90a4688490a468ff90a468ff90a468ff90a468ffa8b788ffe7ecdfff ,
                        0x90a468ff90a4688400000000000000000000000000000000727272ff727272ff ,
                        0x727272e17e866e1790a468cc90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x93a66cd190a46809000000000000000000000000000000000000000000000000 ,
                        0x000000000000000090a4680690a4688190a468d590a468ff90a468d590a46881 ,
                        0x90a4680600000000
                    End

                    LayoutCachedLeft =18960
                    LayoutCachedTop =720
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =1200
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =19320
                    Top =1275
                    Width =1254
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="MarkAsSent"
                    Caption ="  Mark As Sent"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Mark Selected Buyers Order as Sent"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =19320
                    LayoutCachedTop =1275
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =1575
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =19320
                    Top =2100
                    Width =1254
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    Name ="MarkAsDelivered"
                    Caption ="Delivered > D"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Mark Selected Buyers Order as Delivered"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =19320
                    LayoutCachedTop =2100
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =2400
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =4800
                    Top =480
                    Width =13860
                    Height =3900
                    TabIndex =5
                    Name ="BuyersOrderUnits_Sub"
                    SourceObject ="Form.BuyersOrderUnits_Sub"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =480
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =4380
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2955
                    Top =1260
                    Height =237
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    Name ="TakeDelivery"
                    ControlSource ="MustDeliveryDate"
                    Format ="Short Date"
                    ControlTipText ="Set the date the buyer must take delivery"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1497
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1950
                            Top =1260
                            Width =999
                            Height =237
                            FontSize =8
                            FontWeight =700
                            Name ="Label44"
                            Caption ="Take Delivery"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =1260
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =1497
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            BorderThemeColorIndex =2
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =19320
                    Top =1680
                    Width =1254
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    Name ="MarkAsAccepted"
                    Caption ="  Accepted > PS"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Mark Selected Buyers Order as Accepted"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =19320
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =1980
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =19320
                    Top =4020
                    Width =1254
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    Name ="MarkAsDead"
                    Caption ="  Mark Dead > R"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Mark Selected Buyers Order as 'Dead'"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =19320
                    LayoutCachedTop =4020
                    LayoutCachedWidth =20574
                    LayoutCachedHeight =4320
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =3
                    ForeTint =100.0
                    ForeShade =10.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =180
                    Top =600
                    Width =1575
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    ForeColor =855309
                    Name ="NewBuyersOrder"
                    Caption ="  Create New \015\012"
                    OnClick ="[Event Procedure]"
                    FontName ="0ystem"
                    ControlTipText ="Create an Invoice Request"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727296727272f0727272a2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272cc727272ff727272e1 ,
                        0x72727203000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000000000007272722a00000000 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff0000000000000000000000000000000072727218727272337272721b ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff000000000000000000000000000000007272727e727272ff7272728d ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727287727272ff72727296 ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727293727272ff727272a5 ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727299727272ff727272b1 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272a5727272ff727272ba ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272b1727272ff727272c3 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272b7727272ff727272d2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272c3727272ff727272db ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =600
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =900
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeTint =95.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =6
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8210719
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =6
                    PressedTint =60.0
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =855309
                    PressedForeTint =95.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2955
                    Top =1770
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    Name ="TaxRateToUse"
                    ControlSource ="TaxRateToUse"
                    Format ="Percent"
                    ControlTipText ="Set the tax rate to use, default = 0%"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =1770
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =2010
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1950
                            Top =1770
                            Width =999
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label53"
                            Caption ="Tax Rate"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =1770
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =2010
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            BackThemeColorIndex =6
                            BackTint =60.0
                            BorderThemeColorIndex =2
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =16485
                    Top =60
                    Width =990
                    FontSize =8
                    TabIndex =2
                    Name ="BuyerID"
                    ControlSource ="BuyerID"
                    GridlineColor =0

                    LayoutCachedLeft =16485
                    LayoutCachedTop =60
                    LayoutCachedWidth =17475
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =15810
                            Top =75
                            Width =675
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="BuyerIDlbl"
                            Caption ="BuyerID:"
                            GridlineColor =0
                            LayoutCachedLeft =15810
                            LayoutCachedTop =75
                            LayoutCachedWidth =16485
                            LayoutCachedHeight =315
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =50.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =15135
                    Top =60
                    Width =675
                    FontSize =8
                    TabIndex =1
                    Name ="Text58"
                    ControlSource ="InvoiceRequestID"
                    GridlineColor =0

                    LayoutCachedLeft =15135
                    LayoutCachedTop =60
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =14880
                            Top =75
                            Width =255
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label59"
                            Caption ="ID:"
                            GridlineColor =0
                            LayoutCachedLeft =14880
                            LayoutCachedTop =75
                            LayoutCachedWidth =15135
                            LayoutCachedHeight =315
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =50.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =1740
                    Width =8640
                    Height =420
                    FontSize =16
                    FontWeight =700
                    TabIndex =19
                    ForeColor =6108695
                    Name ="Text60"
                    ControlSource ="=\"<<< Create and/or Select a Buyers Order to Add or View Units\""

                    LayoutCachedLeft =4860
                    LayoutCachedTop =1740
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =2160
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18960
                    Top =1275
                    Width =300
                    Height =300
                    FontSize =10
                    TabIndex =15
                    BorderColor =8210719
                    Name ="Text388"
                    ControlSource ="=IIf(IsDate([InvoiceSentDate]),\"✓\",\"X\")"
                    FontName ="Segoe UI"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18960
                    LayoutCachedTop =1275
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =1575
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18960
                    Top =1680
                    Width =300
                    Height =300
                    FontSize =10
                    TabIndex =17
                    BorderColor =8210719
                    Name ="Text61"
                    ControlSource ="=IIf(IsDate([Accepted]),\"✓\",\"X\")"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18960
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =1980
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18960
                    Top =2100
                    Width =300
                    Height =300
                    FontSize =10
                    TabIndex =20
                    BorderColor =8210719
                    Name ="Text62"
                    ControlSource ="=IIf(DMin(\"Status\",\"vw_BuyersOrder\",\"InvoiceRequestID=\" & Nz([InvoiceReque"
                        "stID],0))=\"D\",\"✓\",\"X\")"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18960
                    LayoutCachedTop =2100
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =2400
                    LayoutGroup =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =5
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =19365
                    Top =60
                    Width =990
                    FontSize =8
                    TabIndex =3
                    Name ="RequestedBy"
                    ControlSource ="RequestedBy"
                    GridlineColor =0

                    LayoutCachedLeft =19365
                    LayoutCachedTop =60
                    LayoutCachedWidth =20355
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =18690
                            Top =75
                            Width =675
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label64"
                            Caption ="User:"
                            GridlineColor =0
                            LayoutCachedLeft =18690
                            LayoutCachedTop =75
                            LayoutCachedWidth =19365
                            LayoutCachedHeight =315
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =50.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2955
                    Top =1500
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    Name ="Disposition"
                    ControlSource ="Disposition"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_RMKTDisposition.DispositionID, Lookup_RMKTDisposition.DispositionT"
                        "ype FROM Lookup_RMKTDisposition ORDER BY Lookup_RMKTDisposition.DispositionID;"
                    ColumnWidths ="0;1152"
                    DefaultValue ="=Null"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1770
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =4
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1950
                            Top =1500
                            Width =999
                            Height =270
                            FontSize =8
                            FontWeight =700
                            BorderColor =8210719
                            Name ="Label1806"
                            Caption ="Disposition"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1950
                            LayoutCachedTop =1500
                            LayoutCachedWidth =2949
                            LayoutCachedHeight =1770
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            BackThemeColorIndex =8
                            BackShade =75.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "RemarketingInvoiceRequest_Sub.cls"
