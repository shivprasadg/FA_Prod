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
    Width =25728
    DatasheetFontHeight =11
    ItemSuffix =359
    Left =5295
    Top =6315
    Right =31005
    Bottom =20460
    ShortcutMenuBar ="SortClip"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xc6e830e01372e640
    End
    RecordSource ="SELECT Vendors.VendorSNm AS Vendor, E.InvoiceNumber, E.ExpenseDate, E.Expense, T"
        ".RemarketExpense as ExpenseType, E.Comment, E.InvoiceDate, E.InvoicePaidDate, E."
        "InvoiceAmount, Clients.ClientShNm AS Client, E.AssetID, U.UnitVIN AS VIN, U.Unit"
        "UnitNum AS UnitNum, U.UnitStatus, [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitModel"
        "] & ' ' & [UnitSubType] & ' ' & [UnitType] AS UnitDesc, G.Assignee_Bank AS Assig"
        "nee, RS.ReimburseStatus, E.ReimburseDate, E.RInvoiceNum, BuyOutCost, BuyOutDate "
        ", E.RMKTExpenseId, E.ExpenseTypeId, G.AssigneeBankID, E.ReimburseStatusID, E.Ven"
        "dorId  FROM ((((((RemarketExpense As E  INNER JOIN Units As U ON E.AssetId = U.U"
        "nitID) INNER JOIN Clients ON U.ClientID = Clients.clientID) INNER JOIN SchGrp as"
        " G ON U.GroupREF = G.SGrpID) INNER JOIN LookUp_ReimburseStatus As RS ON E.Reimbu"
        "rseStatusID = RS.ReimburseStatusID) INNER JOIN RemarketExpenseType as T ON E.Exp"
        "enseTypeId = T.RemarketExpenseTypeID) INNER JOIN RMKT As R ON R.UnitRef = U.Unit"
        "ID) LEFT JOIN Vendors ON E.VendorID = Vendors.VendorID  WHERE (((E.ExpenseDate)>"
        "#12/31/2022#)) AND E.InvoiceNumber='63641'  AND ExpenseDate =#8/25/2025#  ORDER "
        "BY InvoiceNumber, E.AssetID, ExpenseDate, ClientShNm"
    Caption ="UnitsIn_sub"
    OnCurrent ="[Event Procedure]"
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
            Height =630
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =420
                    Width =25632
                    Height =210
                    BorderColor =15527148
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =420
                    LayoutCachedWidth =25632
                    LayoutCachedHeight =630
                    TabIndex =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =6360
                    Top =36
                    Width =1092
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitFinalCost_Label"
                    Caption ="Expense\015\012Amount"
                    GroupTable =6
                    LayoutCachedLeft =6360
                    LayoutCachedTop =36
                    LayoutCachedWidth =7452
                    LayoutCachedHeight =492
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =14328
                    Top =36
                    Width =1008
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =6
                    LayoutCachedLeft =14328
                    LayoutCachedTop =36
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =492
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
                    Left =10740
                    Top =36
                    Width =2352
                    Height =456
                    FontSize =9
                    Name ="ULoc_Label"
                    Caption ="Client"
                    GroupTable =6
                    LayoutCachedLeft =10740
                    LayoutCachedTop =36
                    LayoutCachedWidth =13092
                    LayoutCachedHeight =492
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =120
                    Top =36
                    Width =2175
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitPlateNum_Label"
                    Caption ="Vendor"
                    GroupTable =6
                    LayoutCachedLeft =120
                    LayoutCachedTop =36
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =492
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5280
                    Top =36
                    Width =1008
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitEstDelDate_Label"
                    Caption ="Expense\015\012Date"
                    GroupTable =6
                    LayoutCachedLeft =5280
                    LayoutCachedTop =36
                    LayoutCachedWidth =6288
                    LayoutCachedHeight =492
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =15396
                    Top =36
                    Width =1956
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitVIN_Label"
                    Caption ="♦VIN"
                    GroupTable =6
                    LayoutCachedLeft =15396
                    LayoutCachedTop =36
                    LayoutCachedWidth =17352
                    LayoutCachedHeight =492
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =17904
                    Top =36
                    Width =2832
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitYr_Label"
                    Caption ="Unit \015\012Description"
                    GroupTable =6
                    LayoutCachedLeft =17904
                    LayoutCachedTop =36
                    LayoutCachedWidth =20736
                    LayoutCachedHeight =492
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    Left =8964
                    Top =36
                    Width =1716
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label86"
                    Caption ="Comment"
                    GroupTable =6
                    LayoutCachedLeft =8964
                    LayoutCachedTop =36
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =492
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =7524
                    Top =36
                    Width =1368
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label106"
                    Caption ="Expense\015\012Type"
                    GroupTable =6
                    LayoutCachedLeft =7524
                    LayoutCachedTop =36
                    LayoutCachedWidth =8892
                    LayoutCachedHeight =492
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
                    TextAlign =1
                    Left =2364
                    Top =36
                    Width =1680
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="unitinserv_Label"
                    Caption ="Vendor \015\012Invoice"
                    GroupTable =6
                    LayoutCachedLeft =2364
                    LayoutCachedTop =36
                    LayoutCachedWidth =4044
                    LayoutCachedHeight =492
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4116
                    Top =36
                    Width =1092
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitCOASent_Label"
                    Caption ="Vendor\015\012Invoice Date"
                    GroupTable =6
                    LayoutCachedLeft =4116
                    LayoutCachedTop =36
                    LayoutCachedWidth =5208
                    LayoutCachedHeight =492
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =13152
                    Top =36
                    Width =1116
                    Height =456
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label299"
                    Caption ="Assignee"
                    GroupTable =6
                    LayoutCachedLeft =13152
                    LayoutCachedTop =36
                    LayoutCachedWidth =14268
                    LayoutCachedHeight =492
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =25404
                    Top =36
                    Width =288
                    Height =456
                    Name ="EmptyCell307"
                    GroupTable =6
                    LayoutCachedLeft =25404
                    LayoutCachedTop =36
                    LayoutCachedWidth =25692
                    LayoutCachedHeight =492
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =22740
                    Top =36
                    Width =1092
                    Height =456
                    FontSize =9
                    Name ="ReimburseDateLabel"
                    Caption ="R Event \015\012Date"
                    GroupTable =6
                    LayoutCachedLeft =22740
                    LayoutCachedTop =36
                    LayoutCachedWidth =23832
                    LayoutCachedHeight =492
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =20808
                    Top =36
                    Width =1860
                    Height =456
                    FontSize =9
                    Name ="Label310"
                    Caption ="Reimburse \015\012Status"
                    GroupTable =6
                    LayoutCachedLeft =20808
                    LayoutCachedTop =36
                    LayoutCachedWidth =22668
                    LayoutCachedHeight =492
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =23904
                    Top =36
                    Width =1428
                    Height =456
                    FontSize =9
                    Name ="Label346"
                    Caption ="R Invoice \015\012Number"
                    GroupTable =6
                    LayoutCachedLeft =23904
                    LayoutCachedTop =36
                    LayoutCachedWidth =25332
                    LayoutCachedHeight =492
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =17412
                    Top =36
                    Width =432
                    Height =456
                    FontSize =9
                    FontWeight =700
                    Name ="Label354"
                    Caption ="STS"
                    GroupTable =6
                    LayoutCachedLeft =17412
                    LayoutCachedTop =36
                    LayoutCachedWidth =17844
                    LayoutCachedHeight =492
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =25272
                    Top =36
                    Width =288
                    Height =252
                    FontSize =8
                    FontWeight =700
                    Name ="DeleteAllShown"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =7
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

                    LayoutCachedLeft =25272
                    LayoutCachedTop =36
                    LayoutCachedWidth =25560
                    LayoutCachedHeight =288
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =7
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
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
                    Name ="RMKTExpenseId"
                    ControlSource ="RMKTExpenseId"

                    LayoutCachedWidth =645
                    LayoutCachedHeight =255
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8964
                    Top =36
                    Width =1716
                    Height =252
                    FontSize =9
                    TabIndex =7
                    BorderColor =12765388
                    Name ="Comment"
                    ControlSource ="Comment"
                    GroupTable =6

                    LayoutCachedLeft =8964
                    LayoutCachedTop =36
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =288
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
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14328
                    Top =36
                    Width =1008
                    Height =252
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =10
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =14328
                    LayoutCachedTop =36
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =288
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
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Top =36
                    Width =1008
                    Height =252
                    ColumnWidth =2700
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =4
                    BorderColor =12765388
                    Name ="ExpenseDate"
                    ControlSource ="ExpenseDate"
                    Format ="mm/dd/yy"
                    GroupTable =6

                    LayoutCachedLeft =5280
                    LayoutCachedTop =36
                    LayoutCachedWidth =6288
                    LayoutCachedHeight =288
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15396
                    Top =36
                    Width =1956
                    Height =252
                    ColumnWidth =2100
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =11
                    Name ="UnitVIN"
                    ControlSource ="Vin"
                    StatusBarText ="VIN number must be unique in table"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =15396
                    LayoutCachedTop =36
                    LayoutCachedWidth =17352
                    LayoutCachedHeight =288
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17904
                    Top =36
                    Width =2832
                    Height =252
                    ColumnWidth =990
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =13
                    Name ="UnitYr"
                    ControlSource ="UnitDesc"
                    GroupTable =6

                    LayoutCachedLeft =17904
                    LayoutCachedTop =36
                    LayoutCachedWidth =20736
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4116
                    Top =36
                    Width =1092
                    Height =252
                    ColumnWidth =1290
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =3
                    BorderColor =12765388
                    Name ="InvoiceDate"
                    ControlSource ="InvoiceDate"
                    Format ="mm/dd/yy"
                    GroupTable =6

                    LayoutCachedLeft =4116
                    LayoutCachedTop =36
                    LayoutCachedWidth =5208
                    LayoutCachedHeight =288
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
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =36
                    Width =1092
                    Height =252
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =5
                    BorderColor =12765388
                    Name ="UnitFinalCost"
                    ControlSource ="Expense"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =6

                    LayoutCachedLeft =6360
                    LayoutCachedTop =36
                    LayoutCachedWidth =7452
                    LayoutCachedHeight =288
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10740
                    Top =36
                    Width =2352
                    Height =252
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =8
                    Name ="ULoc"
                    ControlSource ="Client"
                    GroupTable =6

                    LayoutCachedLeft =10740
                    LayoutCachedTop =36
                    LayoutCachedWidth =13092
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2364
                    Top =36
                    Width =1680
                    Height =252
                    ColumnWidth =1275
                    FontSize =9
                    TabIndex =2
                    BorderColor =12765388
                    Name ="InvoiceNumber"
                    ControlSource ="InvoiceNumber"
                    GroupTable =6

                    LayoutCachedLeft =2364
                    LayoutCachedTop =36
                    LayoutCachedWidth =4044
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =120
                    Top =36
                    Width =2175
                    Height =252
                    FontSize =9
                    TabIndex =1
                    BorderColor =12765388
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName FROM Vendors INNER JOIN RemarketExpe"
                        "nse ON Vendors.VendorID = RemarketExpense.VendorId GROUP BY Vendors.VendorID, Ve"
                        "ndors.VendorName ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =120
                    LayoutCachedTop =36
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7524
                    Top =36
                    Width =1368
                    Height =252
                    FontSize =9
                    TabIndex =6
                    BorderColor =12765388
                    Name ="ExpenseTypeId"
                    ControlSource ="ExpenseTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketExpenseType.RemarketExpenseTypeId, RemarketExpenseType.RemarketEx"
                        "pense FROM RemarketExpenseType ORDER BY RemarketExpenseType.RemarketExpense;"
                    ColumnWidths ="0;1440"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =7524
                    LayoutCachedTop =36
                    LayoutCachedWidth =8892
                    LayoutCachedHeight =288
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
                    OverlapFlags =127
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =29
                    Top =285
                    Width =25632
                    Height =59
                    FontSize =10
                    TabIndex =18
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0052004d004b00540045007800700065006e0073006500490064005d003d00 ,
                        0x5b00740078007400430075007200720065006e0074005d0000000000
                    End

                    LayoutCachedLeft =29
                    LayoutCachedTop =285
                    LayoutCachedWidth =25661
                    LayoutCachedHeight =344
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d001c0000005b00 ,
                        0x52004d004b00540045007800700065006e0073006500490064005d003d005b00 ,
                        0x740078007400430075007200720065006e0074005d0000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13152
                    Top =36
                    Width =1116
                    Height =252
                    FontSize =9
                    TabIndex =9
                    Name ="Assignee"
                    ControlSource ="Assignee"
                    GroupTable =6

                    LayoutCachedLeft =13152
                    LayoutCachedTop =36
                    LayoutCachedWidth =14268
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =119
                    TextFontCharSet =0
                    Left =25404
                    Top =36
                    Width =288
                    Height =252
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    Name ="DeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =6
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

                    LayoutCachedLeft =25404
                    LayoutCachedTop =36
                    LayoutCachedWidth =25692
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =6
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =22740
                    Top =36
                    Width =1092
                    Height =252
                    FontSize =9
                    TabIndex =15
                    BorderColor =12765388
                    Name ="ReimburseDate"
                    ControlSource ="ReimburseDate"
                    Format ="Short Date"
                    GroupTable =6

                    LayoutCachedLeft =22740
                    LayoutCachedTop =36
                    LayoutCachedWidth =23832
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextFontCharSet =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =20808
                    Top =36
                    Width =1860
                    Height =252
                    FontSize =9
                    TabIndex =14
                    BorderColor =12765388
                    Name ="ReimburseStatusID"
                    ControlSource ="ReimburseStatusID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ReimburseStatus.ReimburseStatusID, LookUp_ReimburseStatus.Reimburs"
                        "eStatus FROM LookUp_ReimburseStatus;"
                    ColumnWidths ="0;1440"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =20808
                    LayoutCachedTop =36
                    LayoutCachedWidth =22668
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23904
                    Top =36
                    Width =1428
                    Height =252
                    FontSize =9
                    TabIndex =16
                    BorderColor =12765388
                    Name ="RInvoiceNum"
                    ControlSource ="RInvoiceNum"
                    GroupTable =6

                    LayoutCachedLeft =23904
                    LayoutCachedTop =36
                    LayoutCachedWidth =25332
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
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
                    Left =17412
                    Top =36
                    Width =432
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =6

                    LayoutCachedLeft =17412
                    LayoutCachedTop =36
                    LayoutCachedWidth =17844
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
            End
        End
        Begin FormFooter
            Height =240
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =240
                    Width =1979
                    FontSize =8
                    BorderColor =12632256
                    Name ="UnitCount"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =2219
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3180
                    Width =765
                    Height =210
                    FontSize =8
                    TabIndex =1
                    BorderColor =12632256
                    Name ="txtCurrent"

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =3945
                    LayoutCachedHeight =210
                End
            End
        End
    End
End
CodeBehindForm
' See "RMKTExpenseSub.cls"
