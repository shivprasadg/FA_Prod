Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    SubdatasheetExpanded = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16275
    DatasheetFontHeight =11
    ItemSuffix =86
    Left =5310
    Top =5340
    Right =17160
    Bottom =14805
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6ad8a0bd2c78e440
    End
    RecordSource ="SELECT RemarketExpense.* FROM RemarketExpense;"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    SplitFormOrientation =1
    SplitFormSize =3015
    SplitFormOrientation =1
    SplitFormSize =3015
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =2100
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =45
                    Top =1020
                    Width =16230
                    Height =360
                    BackColor =8210719
                    BorderColor =10855845
                    Name ="Box22"
                    HorizontalAnchor =2
                    LayoutCachedLeft =45
                    LayoutCachedTop =1020
                    LayoutCachedWidth =16275
                    LayoutCachedHeight =1380
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    Left =13860
                    Width =1950
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =-2147483614
                    Name ="Label33"
                    Caption ="Manage"
                    HorizontalAnchor =1
                    LayoutCachedLeft =13860
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =510
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =14130
                    Top =15
                    Width =1695
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label32"
                    Caption ="Manage"
                    HorizontalAnchor =1
                    LayoutCachedLeft =14130
                    LayoutCachedTop =15
                    LayoutCachedWidth =15825
                    LayoutCachedHeight =525
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =8970
                    Top =165
                    Width =6915
                    Height =1185
                    FontSize =48
                    FontWeight =700
                    ForeColor =15527148
                    Name ="Label23"
                    Caption ="Expenses"
                    HorizontalAnchor =1
                    LayoutCachedLeft =8970
                    LayoutCachedTop =165
                    LayoutCachedWidth =15885
                    LayoutCachedHeight =1350
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =9180
                    Top =180
                    Width =6735
                    Height =1200
                    FontSize =48
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label24"
                    Caption ="Expenses"
                    HorizontalAnchor =1
                    LayoutCachedLeft =9180
                    LayoutCachedTop =180
                    LayoutCachedWidth =15915
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2760
                    Top =90
                    Width =720
                    Height =288
                    TabIndex =2
                    Name ="GoHome"
                    Caption ="GoHome"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000adda0fff0bb0fff0da ,
                        0xad0fff0bb0fff0adda0fff0bb0fff0daad0fff0bb0fff0adda0fff0000fff0da ,
                        0xad0ffffffffff0add0b0ffffffff0b0aad0b0ffffff0b0addad0b0ffff0b0ada ,
                        0xadad0b0ff0b00daddadad0b00b010adaadadad0bb0a10daddadadad00ad10ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="\"Open Main Menu\""
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2760
                    LayoutCachedTop =90
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =378
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1590
                    Top =750
                    Width =990
                    Height =165
                    FontSize =8
                    ForeColor =-2147483614
                    Name ="txtCurrent"

                    LayoutCachedLeft =1590
                    LayoutCachedTop =750
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =915
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =3495
                    Top =90
                    Width =1035
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483617
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3495
                    LayoutCachedTop =90
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =378
                    PictureCaptionArrangement =5
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =48
                    Top =1740
                    Width =1512
                    Height =315
                    FontSize =10
                    Name ="TypeLabel"
                    Caption ="Type:"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =48
                    LayoutCachedTop =1740
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =2055
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5328
                    Top =1740
                    Width =1188
                    Height =315
                    FontSize =10
                    BackColor =15523798
                    BorderColor =-2147483615
                    Name ="Label49"
                    Caption ="Expense Date:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =5328
                    LayoutCachedTop =1740
                    LayoutCachedWidth =6516
                    LayoutCachedHeight =2055
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4164
                    Top =1740
                    Width =1092
                    Height =315
                    FontSize =10
                    BackColor =16317438
                    BorderColor =4754549
                    Name ="Label50"
                    Caption ="Expense:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =4164
                    LayoutCachedTop =1740
                    LayoutCachedWidth =5256
                    LayoutCachedHeight =2055
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9972
                    Top =1740
                    Width =1092
                    Height =315
                    FontSize =10
                    BackColor =16317438
                    BorderColor =4754549
                    Name ="Label51"
                    Caption ="Inv Amount:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =9972
                    LayoutCachedTop =1740
                    LayoutCachedWidth =11064
                    LayoutCachedHeight =2055
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =14988
                    Top =1740
                    Width =768
                    Height =315
                    Name ="Label57"
                    Caption ="ID"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =14988
                    LayoutCachedTop =1740
                    LayoutCachedWidth =15756
                    LayoutCachedHeight =2055
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =6588
                    Top =1740
                    Width =2052
                    Height =315
                    FontSize =10
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="Label68"
                    Caption ="Invoice Number:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =6588
                    LayoutCachedTop =1740
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =2055
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4380
                    Top =1470
                    Width =9525
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BorderColor =4754549
                    ForeColor =-2147483615
                    Name ="Label70"
                    Caption ="┌────────────────────────────────────────────── Edit Here ──────────────────────"
                        "──────────────────────┐"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =1470
                    LayoutCachedWidth =13905
                    LayoutCachedHeight =1725
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =195
                    Top =150
                    Width =705
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="AssetIdLBel"
                    Caption ="AssetId"
                    LayoutCachedLeft =195
                    LayoutCachedTop =150
                    LayoutCachedWidth =900
                    LayoutCachedHeight =435
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =5
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =225
                    Top =345
                    Width =2385
                    Height =420
                    FontSize =18
                    TabIndex =3
                    Name ="AssetId"
                    ControlSource ="AssetId"

                    LayoutCachedLeft =225
                    LayoutCachedTop =345
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =765
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =12396
                    Top =1740
                    Width =2520
                    Height =315
                    FontSize =10
                    BackColor =13434879
                    BorderColor =2838138
                    Name ="Label46"
                    Caption ="Comment"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =12396
                    LayoutCachedTop =1740
                    LayoutCachedWidth =14916
                    LayoutCachedHeight =2055
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =1620
                    Top =1740
                    Width =2475
                    Height =315
                    FontSize =10
                    Name ="Label44"
                    Caption ="Vendor"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =1620
                    LayoutCachedTop =1740
                    LayoutCachedWidth =4095
                    LayoutCachedHeight =2055
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =5625
                    Top =477
                    Width =1470
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =-2147483617
                    Name ="AddNew"
                    Caption ="Add Expense"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5625
                    LayoutCachedTop =477
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =765
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =238
                    Left =3465
                    Top =477
                    Width =2160
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =-2147483617
                    Name ="ChangeMe"
                    Caption ="Change This Expense"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3465
                    LayoutCachedTop =477
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =765
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =7110
                    Top =477
                    Width =2160
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    ForeColor =1643706
                    Name ="DeleteMe"
                    Caption ="Delete This Expense"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7110
                    LayoutCachedTop =477
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =765
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =8712
                    Top =1740
                    Width =1188
                    Height =315
                    FontSize =10
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="Label82"
                    Caption =" Invoice Date:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =8712
                    LayoutCachedTop =1740
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =2055
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =11136
                    Top =1740
                    Width =1188
                    Height =315
                    FontSize =10
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="Label84"
                    Caption ="Inv Paid"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =11136
                    LayoutCachedTop =1740
                    LayoutCachedWidth =12324
                    LayoutCachedHeight =2055
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =9285
                    Top =477
                    Width =1470
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    ForeColor =-2147483617
                    Name ="SaveUpdate"
                    Caption ="Save | Update"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9285
                    LayoutCachedTop =477
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =765
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7380
            Name ="Detail"
            AlternateBackColor =15527148
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =4164
                    Top =48
                    Width =1092
                    Height =312
                    FontSize =10
                    TabIndex =2
                    BackColor =15858167
                    BorderColor =4754549
                    Name ="Expense"
                    ControlSource ="Expense"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =4164
                    LayoutCachedTop =48
                    LayoutCachedWidth =5256
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =9972
                    Top =48
                    Width =1092
                    Height =312
                    FontSize =10
                    TabIndex =6
                    BackColor =15858167
                    BorderColor =4754549
                    Name ="InvoiceAmount"
                    ControlSource ="InvoiceAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9972
                    LayoutCachedTop =48
                    LayoutCachedWidth =11064
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =765
                    Width =15840
                    Height =29
                    TabIndex =10
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000096000000010000000100000000000000000000001a00000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00570061007200720061006e0074007900490064005d003d005b0074007800 ,
                        0x7400430075007200720065006e0074005d0000000000
                    End

                    LayoutCachedTop =765
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =794
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00190000005b00 ,
                        0x570061007200720061006e0074007900490064005d003d005b00740078007400 ,
                        0x430075007200720065006e0074005d0000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5328
                    Top =48
                    Width =1188
                    Height =312
                    FontSize =10
                    TabIndex =3
                    BackColor =15523798
                    BorderColor =-2147483615
                    Name ="ExpenseDate"
                    ControlSource ="ExpenseDate"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5328
                    LayoutCachedTop =48
                    LayoutCachedWidth =6516
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =6588
                    Top =48
                    Width =2052
                    Height =312
                    FontSize =10
                    TabIndex =4
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="InvoiceNumber"
                    ControlSource ="InvoiceNumber"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6588
                    LayoutCachedTop =48
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =12
                    Left =225
                    Top =1035
                    Width =15840
                    Height =4545
                    FontSize =10
                    TabIndex =11
                    ForeColor =8210719
                    Name ="ExpenseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketExpense.RMKTExpenseId, RemarketExpense.AssetId, RemarketExpenseTy"
                        "pe.RemarketExpense AS Type, Vendors.VendorSNm, RemarketExpense.ExpenseDate, Rema"
                        "rketExpense.Expense, RemarketExpense.InvoiceNumber, RemarketExpense.InvoiceDate,"
                        " RemarketExpense.InvoiceAmount, RemarketExpense.InvoicePaidDate, RemarketExpense"
                        ".Comment FROM (RemarketExpense LEFT JOIN Vendors ON RemarketExpense.VendorId = V"
                        "endors.VendorID) LEFT JOIN RemarketExpenseType ON RemarketExpense.ExpenseTypeId "
                        "= RemarketExpenseType.RemarketExpenseTypeId WHERE (((RemarketExpense.AssetId)=[F"
                        "orms]![Remarketing_Costs]![RemarketExpense].[Form]![AssetId]));"
                    ColumnWidths ="0;0;1296;1728;1152;1152;1440;1440;1152;1728;4176"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select Warranty to Edit"

                    LayoutCachedLeft =225
                    LayoutCachedTop =1035
                    LayoutCachedWidth =16065
                    LayoutCachedHeight =5580
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =12396
                    Top =48
                    Width =2520
                    Height =312
                    FontSize =10
                    TabIndex =9
                    BackColor =13434879
                    BorderColor =2838138
                    Name ="Comment"
                    ControlSource ="Comment"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12396
                    LayoutCachedTop =48
                    LayoutCachedWidth =14916
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =5
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14988
                    Top =48
                    Width =768
                    Height =312
                    TabIndex =8
                    Name ="RMKTExpenseId"
                    ControlSource ="RMKTExpenseId"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =14988
                    LayoutCachedTop =48
                    LayoutCachedWidth =15756
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =8712
                    Top =48
                    Width =1188
                    Height =312
                    FontSize =10
                    TabIndex =5
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="Text81"
                    ControlSource ="InvoiceDate"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8712
                    LayoutCachedTop =48
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =11136
                    Top =48
                    Width =1188
                    Height =312
                    FontSize =10
                    TabIndex =7
                    BackColor =13952764
                    BorderColor =2838138
                    Name ="Text83"
                    ControlSource ="InvoicePaidDate"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =11136
                    LayoutCachedTop =48
                    LayoutCachedWidth =12324
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =5
                    OverlapFlags =93
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1620
                    Top =48
                    Width =2475
                    Height =312
                    FontSize =10
                    TabIndex =1
                    BoundColumn =1
                    Name ="VendorId"
                    ControlSource ="VendorId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorName, Vendors.VendorID FROM Vendors;"
                    ColumnWidths ="2592;432"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =1620
                    LayoutCachedTop =48
                    LayoutCachedWidth =4095
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    SpecialEffect =5
                    OverlapFlags =93
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =48
                    Top =48
                    Width =1512
                    Height =312
                    FontSize =10
                    BoundColumn =1
                    Name ="ExpenseTypeId"
                    ControlSource ="ExpenseTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketExpenseType.RemarketExpense AS Expense, RemarketExpenseType.Remar"
                        "ketExpenseTypeId AS ID FROM RemarketExpenseType;"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =48
                    LayoutCachedTop =48
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =2
                    Left =165
                    Top =360
                    Width =3645
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =4754549
                    ForeColor =-2147483615
                    Name ="EditHereAlso"
                    Caption ="└──────────── Edit Here Also─────────────┘"
                    LayoutCachedLeft =165
                    LayoutCachedTop =360
                    LayoutCachedWidth =3810
                    LayoutCachedHeight =585
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
' See "RemarketExpense.cls"
