Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    TabularCharSet =162
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14460
    DatasheetFontHeight =11
    ItemSuffix =1617
    Left =5130
    Top =13170
    Right =18165
    Bottom =19965
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x2bad22cd72d9e540
    End
    RecordSource ="SELECT Rmkt.* FROM Rmkt WHERE (((Rmkt.UnitRef)=[Forms]![UnitsForm]![UnitID]));"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            BorderColor =12632256
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =7512
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =5490
                    Top =2445
                    Width =1290
                    Height =270
                    FontSize =10
                    Name ="Label1806"
                    Caption ="Disposition"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5490
                    LayoutCachedTop =2445
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =2715
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    Width =7575
                    Height =465
                    FontSize =18
                    Name ="Label690"
                    Caption ="Remarketing Details for this Unit"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =7575
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =7
                    ForeShade =50.0
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Top =3480
                    Width =13140
                    Height =4032
                    TabIndex =37
                    Name ="RmktExpense"
                    SourceObject ="Form.RmktExpense"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =3480
                    LayoutCachedWidth =13140
                    LayoutCachedHeight =7512
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3780
                    Top =1440
                    Width =1590
                    Height =270
                    TabIndex =10
                    BackColor =13434879
                    BorderColor =3751056
                    Name ="BuyDateEffective"
                    ControlSource ="BuyDateEffective"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date Buy-Out was Effective; normally equals Actaul Offlease Date"
                    OnDirty ="[Event Procedure]"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3780
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =1710
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2895
                            Top =1440
                            Width =825
                            Height =285
                            FontSize =10
                            BackColor =13434879
                            Name ="Label10"
                            Caption =" Effective"
                            TopPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2895
                            LayoutCachedTop =1440
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1725
                            ColumnStart =5
                            ColumnEnd =5
                            ForeThemeColorIndex =0
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Top =1125
                    Width =1245
                    Height =270
                    TabIndex =26
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Sale price must be >=0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =1125
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =1395
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =1125
                            Width =968
                            Height =270
                            FontSize =10
                            Name ="Label5"
                            Caption ="Sale Price"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =1125
                            LayoutCachedWidth =9128
                            LayoutCachedHeight =1395
                            RowStart =1
                            RowEnd =1
                            ColumnStart =9
                            ColumnEnd =9
                            BackThemeColorIndex =6
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11805
                    Top =1125
                    Width =1170
                    Height =270
                    TabIndex =33
                    ForeColor =-2147483600
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =11805
                    LayoutCachedTop =1125
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =1395
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10590
                            Top =1125
                            Width =1155
                            Height =270
                            FontSize =10
                            ForeColor =-2147483600
                            Name ="Label6"
                            Caption ="Invoice Sent "
                            LayoutCachedLeft =10590
                            LayoutCachedTop =1125
                            LayoutCachedWidth =11745
                            LayoutCachedHeight =1395
                            RowStart =1
                            RowEnd =1
                            ColumnStart =11
                            ColumnEnd =11
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11805
                    Top =1455
                    Width =1170
                    Height =270
                    TabIndex =34
                    ForeColor =-2147483600
                    Name ="InvoicePaidDate"
                    ControlSource ="InvoicePaidDate"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =11805
                    LayoutCachedTop =1455
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =1725
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =10590
                            Top =1455
                            Width =1155
                            Height =270
                            FontSize =10
                            Name ="Label7"
                            Caption =" Invoice Paid "
                            LayoutCachedLeft =10590
                            LayoutCachedTop =1455
                            LayoutCachedWidth =11745
                            LayoutCachedHeight =1725
                            RowStart =2
                            RowEnd =2
                            ColumnStart =11
                            ColumnEnd =11
                            BackThemeColorIndex =2
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =7200
                    Left =9057
                    Top =2445
                    Width =1748
                    Height =270
                    FontSize =9
                    TabIndex =30
                    Name ="BuyerRef"
                    ControlSource ="BuyerRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Buyers.BuyerId, Buyers.BuyerShortName, Buyers.BuyerName, IIf([Wholeseller"
                        "]=-1,\"W\",\"R\") AS Dspn FROM Buyers ORDER BY Buyers.BuyerShortName;"
                    ColumnWidths ="0;2160;4320;288"
                    StatusBarText ="Link to Buyer"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="This is the Buyer, Double Click to view"

                    LayoutCachedLeft =9057
                    LayoutCachedTop =2445
                    LayoutCachedWidth =10805
                    LayoutCachedHeight =2715
                    DisplayAsHyperlink =1
                    RowStart =5
                    RowEnd =5
                    ColumnStart =10
                    ColumnEnd =12
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =2445
                            Width =846
                            Height =270
                            FontSize =10
                            Name ="Label4"
                            Caption ="♦Buyer"
                            ControlTipText ="Customers are listed after buyers"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =2445
                            LayoutCachedWidth =9006
                            LayoutCachedHeight =2715
                            RowStart =5
                            RowEnd =5
                            ColumnStart =9
                            ColumnEnd =9
                            BackThemeColorIndex =6
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1290
                    Top =1455
                    Width =1155
                    Height =270
                    TabIndex =2
                    BackColor =13233663
                    BorderColor =13233663
                    ForeColor =26265
                    Name ="PakrDt"
                    ControlSource ="ParkDT"
                    Format ="Short Date"
                    StatusBarText ="date unit parked"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1455
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1725
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1455
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label331"
                            Caption ="Parked"
                            LeftPadding =0
                            LayoutCachedTop =1455
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1725
                            RowStart =2
                            RowEnd =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1290
                    Top =825
                    Width =1155
                    Height =270
                    BackColor =13233663
                    BorderColor =13233663
                    ForeColor =26265
                    Name ="SUF Sent"
                    ControlSource ="SUF Sent"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    EventProcPrefix ="SUF_Sent"
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =825
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1095
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =825
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label22"
                            Caption ="SUF Sent"
                            LeftPadding =0
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedTop =825
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1095
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1290
                    Top =1125
                    Width =1155
                    Height =270
                    TabIndex =1
                    BackColor =13233663
                    BorderColor =13233663
                    ForeColor =26265
                    Name ="SUF Returned"
                    ControlSource ="SUF Returned"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    EventProcPrefix ="SUF_Returned"
                    RightPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1125
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1395
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1125
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label23"
                            Caption ="SUF Returned"
                            LeftPadding =0
                            LayoutCachedTop =1125
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1395
                            RowStart =1
                            RowEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1290
                    Top =2445
                    Width =1155
                    Height =270
                    FontSize =10
                    TabIndex =5
                    BorderColor =26265
                    Name ="BankNotified"
                    ControlSource ="BankNotified"
                    Format ="Short Date"
                    RightPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =2445
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =2715
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =525
                            Top =2445
                            Width =705
                            Height =270
                            FontSize =10
                            BackColor =4815819
                            Name ="Label24"
                            Caption ="Bank Notified"
                            LeftPadding =0
                            LayoutCachedLeft =525
                            LayoutCachedTop =2445
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =2715
                            RowStart =5
                            RowEnd =5
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5760
                    Top =120
                    Width =360
                    Height =195
                    FontSize =8
                    TabIndex =40
                    Name ="RmktId"
                    ControlSource ="RmktId"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =120
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =315
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1290
                    Top =1785
                    Width =1155
                    Height =270
                    TabIndex =3
                    BackColor =13233663
                    BorderColor =13233663
                    ForeColor =26265
                    Name ="MiParked"
                    ControlSource ="MiParked"
                    Format ="Standard"
                    StatusBarText ="mileage SUF  when Parked"
                    ValidationRule =">=0"
                    ValidationText ="Mileage must be >=0"
                    DefaultValue ="0"
                    RightPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1785
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =2055
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1785
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label28"
                            Caption ="Mileage"
                            LeftPadding =0
                            LayoutCachedTop =1785
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =2055
                            RowStart =3
                            RowEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1290
                    Top =2115
                    Width =1155
                    Height =270
                    TabIndex =4
                    BackColor =13233663
                    BorderColor =13233663
                    ForeColor =26265
                    Name ="HrParked"
                    ControlSource ="HrParked"
                    Format ="Standard"
                    StatusBarText ="Hours on SUFwhen Parked"
                    ValidationRule =">=0"
                    ValidationText ="Hours must be >=0"
                    DefaultValue ="0"
                    RightPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =2115
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =2385
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =2115
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label29"
                            Caption ="Hours"
                            LeftPadding =0
                            LayoutCachedTop =2115
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =2385
                            RowStart =4
                            RowEnd =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Top =825
                    Width =1245
                    Height =270
                    TabIndex =25
                    Name ="RmktSPRAdj"
                    ControlSource ="RmktSPRAdj"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="if giving money back should be negative"
                    ValidationRule ="Is Not Null"
                    ValidationText ="the Sales Price adjustment field must be 0 or a +/- adjustment amount"
                    DefaultValue ="0"
                    ControlTipText ="if giving money back should be negative"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9240
                    LayoutCachedTop =825
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =1095
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =825
                            Width =1020
                            Height =270
                            FontSize =10
                            Name ="Label30"
                            Caption ="Price Adj"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =8160
                            LayoutCachedTop =825
                            LayoutCachedWidth =9180
                            LayoutCachedHeight =1095
                            ColumnStart =9
                            ColumnEnd =9
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8403
                    Top =2775
                    Width =2402
                    Height =660
                    FontSize =9
                    TabIndex =31
                    Name ="RmktAdjNote"
                    ControlSource ="RmktAdjNote"

                    LayoutCachedLeft =8403
                    LayoutCachedTop =2775
                    LayoutCachedWidth =10805
                    LayoutCachedHeight =3435
                    RowStart =6
                    RowEnd =7
                    ColumnStart =10
                    ColumnEnd =12
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =1125
                    Width =1215
                    Height =270
                    TabIndex =18
                    Name ="TargetSale"
                    ControlSource ="TargetSale"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="price Rmk Sm should get for unit"
                    RightPadding =105

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1125
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =1395
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5490
                            Top =1125
                            Width =1290
                            Height =270
                            FontSize =10
                            ForeColor =39168
                            Name ="Label52"
                            Caption ="Target Price"
                            LeftPadding =90
                            LayoutCachedLeft =5490
                            LayoutCachedTop =1125
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =1395
                            RowStart =3
                            RowEnd =3
                            ColumnStart =7
                            ColumnEnd =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Top =1455
                    Width =1245
                    Height =270
                    TabIndex =27
                    Name ="RmktDeposit"
                    ControlSource ="RmktDeposit"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =1455
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =1725
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5490
                            Top =1455
                            Width =1020
                            Height =270
                            FontSize =10
                            Name ="Label65"
                            Caption ="Deposit"
                            LayoutCachedLeft =5490
                            LayoutCachedTop =1455
                            LayoutCachedWidth =6510
                            LayoutCachedHeight =1725
                            RowStart =2
                            RowEnd =2
                            ColumnStart =9
                            ColumnEnd =9
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Top =1785
                    Width =1245
                    Height =270
                    TabIndex =28
                    Name ="RmktDepdt"
                    ControlSource ="RmktDepdt"
                    Format ="Short Date"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =1785
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =2055
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =1785
                            Width =968
                            Height =270
                            FontSize =10
                            Name ="Label67"
                            Caption ="Received"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =1785
                            LayoutCachedWidth =9128
                            LayoutCachedHeight =2055
                            RowStart =3
                            RowEnd =3
                            ColumnStart =9
                            ColumnEnd =9
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =1785
                    Width =1215
                    Height =270
                    TabIndex =20
                    Name ="FABrkComm"
                    ControlSource ="FABrkComm"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="FA per unit commission on brokered deal"
                    RightPadding =105
                    BottomPadding =165

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1785
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =2055
                    RowStart =6
                    RowEnd =6
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5490
                            Top =1785
                            Width =1290
                            Height =270
                            FontSize =10
                            ForeColor =39168
                            Name ="Label69"
                            Caption ="Broker Comm"
                            LeftPadding =90
                            BottomPadding =165
                            LayoutCachedLeft =5490
                            LayoutCachedTop =1785
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =2055
                            RowStart =6
                            RowEnd =6
                            ColumnStart =7
                            ColumnEnd =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6840
                    Top =825
                    Width =1215
                    Height =270
                    TabIndex =17
                    Name ="AdjBookValue"
                    ControlSource ="AdjBookValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="adjusted book value inclusive of LO support"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =105

                    LayoutCachedLeft =6840
                    LayoutCachedTop =825
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =1095
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =5490
                            Top =825
                            Width =1290
                            Height =270
                            FontSize =10
                            Name ="Label71"
                            Caption ="FA Book Value"
                            LeftPadding =90
                            LayoutCachedLeft =5490
                            LayoutCachedTop =825
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =1095
                            RowStart =1
                            RowEnd =1
                            ColumnStart =7
                            ColumnEnd =7
                            BackThemeColorIndex =6
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Top =465
                    Width =2430
                    Height =239
                    FontSize =8
                    BackColor =4815819
                    Name ="ACUPDT1"
                    Caption ="Parked SUF"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =465
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =704
                    ColumnEnd =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11805
                    Top =1785
                    Width =1170
                    Height =270
                    TabIndex =35
                    ForeColor =-2147483600
                    Name ="RmktBOSSent1"
                    ControlSource ="RmktBOSSent"
                    Format ="Short Date"
                    StatusBarText ="date Accounting sent bill of sale for unit"
                    ShortcutMenuBar ="Enter Title Sent date on Title Log!"
                    ControlTipText ="Enter Title Sent date on Title Log!"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="UpdtBOSsentdtf"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Action ="MoveSize"
                            Argument ="5760"
                            Argument ="10080"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"RmktBOSSent1\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argum"
                                "ent Name=\"FormName\">UpdtBOSsent"
                        End
                        Begin
                            Comment ="_AXL:dtf</Argument></Action><Action Name=\"MoveAndSizeWindow\"><Argument Name=\""
                                "Right\">5760</Argument><Argument Name=\"Down\">10080</Argument></Action></Statem"
                                "ents></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =11805
                    LayoutCachedTop =1785
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =2055
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =10590
                            Top =1785
                            Width =1155
                            Height =270
                            FontSize =10
                            Name ="Label75"
                            Caption ="BOS Sent "
                            ControlTipText ="BOS -Bill of Sales"
                            LayoutCachedLeft =10590
                            LayoutCachedTop =1785
                            LayoutCachedWidth =11745
                            LayoutCachedHeight =2055
                            RowStart =3
                            RowEnd =3
                            ColumnStart =11
                            ColumnEnd =11
                            BackThemeColorIndex =2
                            BackTint =60.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =2115
                    Width =1245
                    Height =270
                    TabIndex =29
                    Name ="RmkSm"
                    ControlSource ="RmkSm"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.OBSOLETE, Emps.Rmkt FROM Emps WHERE (((Emps.OBSOLETE)=N"
                        "o) AND ((Emps.Rmkt)=Yes)) ORDER BY Emps.EmpInt;"
                    StatusBarText ="initials of Remarketing Salesperson"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =2115
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =2385
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8160
                            Top =2115
                            Width =968
                            Height =270
                            FontSize =10
                            Name ="Label32"
                            Caption ="Sales REP"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =2115
                            LayoutCachedWidth =9128
                            LayoutCachedHeight =2385
                            RowStart =4
                            RowEnd =4
                            ColumnStart =9
                            ColumnEnd =9
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2505
                    Top =465
                    Width =2865
                    Height =255
                    FontSize =8
                    BackColor =2185921
                    Name ="Label429"
                    Caption =" Buy Out - INTERIM Funding "
                    FontName ="Century Gothic"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2505
                    LayoutCachedTop =465
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =720
                    ColumnStart =5
                    ColumnEnd =6
                    BorderThemeColorIndex =5
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =5490
                    Top =465
                    Width =2565
                    Height =239
                    FontSize =8
                    Name ="Label431"
                    Caption =" Unit Pricing"
                    FontName ="Century Gothic"
                    LeftPadding =90
                    TopPadding =0
                    RightPadding =105
                    BottomPadding =0
                    LayoutCachedLeft =5490
                    LayoutCachedTop =465
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =704
                    ColumnStart =7
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =8160
                    Top =465
                    Width =2325
                    Height =239
                    FontSize =8
                    Name ="Label434"
                    Caption ="Unit Sold"
                    FontName ="Century Gothic"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8160
                    LayoutCachedTop =465
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =704
                    ColumnStart =9
                    ColumnEnd =10
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =10545
                    Top =465
                    Width =2445
                    Height =239
                    FontSize =8
                    BackColor =12349952
                    Name ="Label438"
                    Caption ="Invoicing"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10545
                    LayoutCachedTop =465
                    LayoutCachedWidth =12990
                    LayoutCachedHeight =704
                    ColumnStart =11
                    ColumnEnd =12
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =1125
                    Width =1590
                    Height =270
                    TabIndex =9
                    BorderColor =3751056
                    Name ="BuyOutCost"
                    ControlSource ="BuyOutCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Amount Wired / Paid"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =1125
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =1395
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2895
                            Top =1125
                            Width =825
                            Height =285
                            FontSize =10
                            BackColor =26316
                            ForeColor =16777215
                            Name ="Label551"
                            Caption =" Cost"
                            GridlineColor =10921638
                            LayoutCachedLeft =2895
                            LayoutCachedTop =1125
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1410
                            RowStart =1
                            RowEnd =1
                            ColumnStart =5
                            ColumnEnd =5
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3780
                    Top =2760
                    Width =1590
                    Height =270
                    TabIndex =14
                    Name ="IFDate"
                    ControlSource ="IFDate"
                    Format ="Short Date"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =2760
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =3030
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =5
                    BackTint =20.0
                    BorderThemeColorIndex =5
                    BorderShade =75.0
                    ForeThemeColorIndex =5
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2520
                            Top =2760
                            Width =1200
                            Height =270
                            FontSize =10
                            Name ="Label561"
                            Caption ="Interim Date"
                            LayoutCachedLeft =2520
                            LayoutCachedTop =2760
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =3030
                            RowStart =4
                            RowEnd =4
                            ColumnStart =5
                            ColumnEnd =5
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3780
                    Top =2100
                    Width =1590
                    Height =270
                    TabIndex =12
                    Name ="IFAmount"
                    ControlSource ="IFAmount"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =2370
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =5
                    BackTint =20.0
                    BorderThemeColorIndex =5
                    BorderShade =75.0
                    ForeThemeColorIndex =5
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2520
                            Top =2100
                            Width =1200
                            Height =270
                            FontSize =10
                            Name ="Label563"
                            Caption ="Interim Amt"
                            LayoutCachedLeft =2520
                            LayoutCachedTop =2100
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =2370
                            RowStart =2
                            RowEnd =2
                            ColumnStart =5
                            ColumnEnd =5
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3780
                    Top =2430
                    Width =1590
                    Height =270
                    TabIndex =13
                    Name ="IFBank"
                    ControlSource ="=IIf(IsNull([IFBankID]),Null,DLookUp(\"BankShortName\",\"[Banks]\",\"[BankId]=\""
                        " & [IFBankID] & \"\"))"
                    ControlTipText ="Interim Funding Bank for the remarketed unit"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =2430
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =2700
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =5
                    BackTint =20.0
                    BorderThemeColorIndex =5
                    BorderShade =75.0
                    ForeThemeColorIndex =5
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2520
                            Top =2430
                            Width =1200
                            Height =270
                            FontSize =10
                            Name ="Label557"
                            Caption ="Interim Bank"
                            LayoutCachedLeft =2520
                            LayoutCachedTop =2430
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =2700
                            RowStart =3
                            RowEnd =3
                            ColumnStart =5
                            ColumnEnd =5
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3780
                    Top =825
                    Width =1590
                    Height =270
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =3751056
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    ControlTipText ="Date Buy-Out was Wired / Paid"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =825
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =1095
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2895
                            Top =810
                            Width =825
                            Height =285
                            FontSize =10
                            BackColor =13434879
                            Name ="Label728"
                            Caption =" Wired"
                            GridlineColor =10921638
                            LayoutCachedLeft =2895
                            LayoutCachedTop =810
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1095
                            RowStart =5
                            RowEnd =5
                            ColumnStart =5
                            ColumnEnd =5
                            ForeThemeColorIndex =0
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6480
                    Top =120
                    Width =360
                    Height =195
                    FontSize =8
                    TabIndex =16
                    Name ="IFBankID"
                    ControlSource ="IFBankID"
                    RightPadding =105

                    LayoutCachedLeft =6480
                    LayoutCachedTop =120
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =315
                    RowStart =7
                    RowEnd =7
                    ColumnStart =8
                    ColumnEnd =8
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =360
                    Height =195
                    FontSize =8
                    TabIndex =41
                    Name ="UnitRef"
                    ControlSource ="UnitRef"
                    LeftPadding =90

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =315
                    RowStart =7
                    RowEnd =7
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11400
                    Top =60
                    Width =1590
                    Height =270
                    FontSize =10
                    TabIndex =39
                    Name ="SaveMe"
                    Caption ="Update / Save "
                    OnClick ="[Event Procedure]"
                    BottomPadding =165

                    LayoutCachedLeft =11400
                    LayoutCachedTop =60
                    LayoutCachedWidth =12990
                    LayoutCachedHeight =330
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    Vertical = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    Left =8160
                    Top =2775
                    Width =248
                    Height =644
                    FontSize =10
                    Name ="Label31"
                    Caption ="Notes"
                    LayoutCachedLeft =8160
                    LayoutCachedTop =2775
                    LayoutCachedWidth =8408
                    LayoutCachedHeight =3419
                    RowStart =6
                    RowEnd =7
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
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
                    Left =6840
                    Top =1455
                    Width =1215
                    Height =270
                    TabIndex =19
                    Name ="GainLoss"
                    ControlSource ="GainLoss"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="price Rmk Sm should get for unit"
                    ControlTipText ="Actual Termanation Penalty, after unit was sold."
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ffcccc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    RightPadding =105

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1455
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =1725
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ffcccc00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =5490
                            Top =1455
                            Width =1290
                            Height =270
                            FontSize =10
                            Name ="TermPenaltyLabel"
                            Caption ="Gain/Loss"
                            LeftPadding =90
                            LayoutCachedLeft =5490
                            LayoutCachedTop =1455
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =1725
                            RowStart =4
                            RowEnd =4
                            ColumnStart =7
                            ColumnEnd =7
                            BackThemeColorIndex =6
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1290
                    Top =3105
                    Width =1155
                    Height =270
                    TabIndex =7
                    BackColor =13233663
                    ForeColor =26265
                    Name ="ParkDateProjected"
                    ControlSource ="ParkDateProjected"
                    Format ="Short Date"
                    StatusBarText ="date unit is PROJECTED to park"
                    RightPadding =0
                    BottomPadding =165

                    LayoutCachedLeft =1290
                    LayoutCachedTop =3105
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =3375
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Top =3105
                            Width =1230
                            Height =270
                            FontSize =10
                            ForeColor =26265
                            Name ="Label1418"
                            Caption ="Proj'd Parked"
                            LeftPadding =0
                            BottomPadding =165
                            LayoutCachedTop =3105
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =3375
                            RowStart =6
                            RowEnd =6
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1005
                    Top =2775
                    Height =270
                    TabIndex =6
                    BackColor =13233663
                    ForeColor =26265
                    Name ="MileageProjected"
                    ControlSource ="MileageProjected"
                    StatusBarText ="The Projected Mileage @ the Projected Park Date"
                    BottomPadding =165

                    LayoutCachedLeft =1005
                    LayoutCachedTop =2775
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =3045
                    RowStart =6
                    RowEnd =6
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =15
                    Top =2790
                    Width =945
                    Height =285
                    FontSize =10
                    ForeColor =26265
                    Name ="Label1433"
                    Caption ="Proj'd Mi"
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =165
                    LayoutCachedLeft =15
                    LayoutCachedTop =2790
                    LayoutCachedWidth =960
                    LayoutCachedHeight =3075
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =3105
                    Width =1590
                    Height =270
                    TabIndex =15
                    BorderColor =3751056
                    Name ="TermPenalty"
                    ControlSource ="TermPenalty"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Protected"
                    ControlTipText ="Projected Loss (@LPAF) to remarket this unit. ONCE ENTERED, THIS VALUE WILL NOT "
                        "CHANGE."
                    BottomPadding =165

                    LayoutCachedLeft =3780
                    LayoutCachedTop =3105
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =3375
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =5
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =2520
                            Top =3105
                            Width =1200
                            Height =270
                            FontSize =8
                            BackColor =2185921
                            Name ="Label1489"
                            Caption ="Term. Penalty"
                            FontName ="Century Gothic"
                            BottomPadding =165
                            LayoutCachedLeft =2520
                            LayoutCachedTop =3105
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =3375
                            RowStart =6
                            RowEnd =6
                            ColumnStart =5
                            ColumnEnd =5
                            BorderThemeColorIndex =5
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11805
                    Top =825
                    Width =1170
                    Height =270
                    TabIndex =32
                    ForeColor =-2147483600
                    Name ="InvoiceRequestDate"
                    ControlSource ="InvoiceRequestDate"
                    StatusBarText ="date invoice was requested from Accounting"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11805
                    LayoutCachedTop =825
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =1095
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10590
                            Top =825
                            Width =1155
                            Height =270
                            FontSize =10
                            ForeColor =-2147483600
                            Name ="Label1600"
                            Caption ="Request Date"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10590
                            LayoutCachedTop =825
                            LayoutCachedWidth =11745
                            LayoutCachedHeight =1095
                            ColumnStart =11
                            ColumnEnd =11
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =2115
                    Width =1215
                    Height =270
                    TabIndex =21
                    Name ="BankShareNetProceeds"
                    ControlSource ="BankShareNetProceeds"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Bank Share Net Proceeds From Sale"
                    RightPadding =105
                    BottomPadding =165

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2115
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =2385
                    RowStart =6
                    RowEnd =6
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5430
                            Top =2115
                            Width =1350
                            Height =270
                            FontSize =10
                            ForeColor =39168
                            Name ="Label1602"
                            Caption ="Bank Share Net "
                            LeftPadding =90
                            BottomPadding =165
                            LayoutCachedLeft =5430
                            LayoutCachedTop =2115
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =2385
                            RowStart =6
                            RowEnd =6
                            ColumnStart =7
                            ColumnEnd =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6840
                    Top =2460
                    Width =1198
                    Height =270
                    FontSize =9
                    TabIndex =22
                    Name ="Wholesale"
                    ControlSource ="Wholesale"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_RMKTDisposition.DispositionID, Lookup_RMKTDisposition.DispositionT"
                        "ype FROM Lookup_RMKTDisposition ORDER BY Lookup_RMKTDisposition.DispositionID;"
                    ColumnWidths ="0;1152"
                    DefaultValue ="=Null"
                    Format ="True/False"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2460
                    LayoutCachedWidth =8038
                    LayoutCachedHeight =2730
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =1770
                    Width =1590
                    Height =270
                    TabIndex =11
                    BorderColor =3751056
                    Name ="BuyOutCostNet"
                    ControlSource ="BuyOutCostNet"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Cost Net = (Amount wired - PD Rents Collected) [Calculated]"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =1770
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =2040
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2895
                            Top =1770
                            Width =825
                            Height =285
                            FontSize =10
                            BackColor =26316
                            ForeColor =16777215
                            Name ="Label1606"
                            Caption ="NET  Cost"
                            GridlineColor =10921638
                            LayoutCachedLeft =2895
                            LayoutCachedTop =1770
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =2055
                            RowStart =2
                            RowEnd =2
                            ColumnStart =5
                            ColumnEnd =5
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =10863
                    Top =2453
                    Width =2099
                    Height =982
                    FontSize =9
                    TabIndex =36
                    ForeColor =6108695
                    BorderColor =6108695
                    Name ="BuyOrdersList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_BuyersOrder.InvoiceRequestID, IIf([Won]=-1,\"Y\",\"-\") AS Accepted, v"
                        "w_BuyersOrder.BuyerShortName, vw_BuyersOrder.BuyerID FROM UnitsFormFilter INNER "
                        "JOIN vw_BuyersOrder ON UnitsFormFilter.AssetIDx = vw_BuyersOrder.UnitID;"
                    ColumnWidths ="504;360;288;0"
                    StatusBarText ="date invoice was requested from Accounting"
                    OnDblClick ="[Event Procedure]"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10863
                    LayoutCachedTop =2453
                    LayoutCachedWidth =12962
                    LayoutCachedHeight =3435
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =10868
                    Top =2115
                    Width =2094
                    Height =275
                    FontSize =10
                    Name ="Label1610"
                    Caption ="Buy Orders "
                    LayoutCachedLeft =10868
                    LayoutCachedTop =2115
                    LayoutCachedWidth =12962
                    LayoutCachedHeight =2390
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6857
                    Top =3115
                    Width =1170
                    Height =270
                    TabIndex =24
                    ForeColor =-2147483600
                    Name ="PayOffReceived"
                    ControlSource ="PayOffReceived"
                    Format ="Short Date"

                    LayoutCachedLeft =6857
                    LayoutCachedTop =3115
                    LayoutCachedWidth =8027
                    LayoutCachedHeight =3385
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =7
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =5490
                            Top =3115
                            Width =1290
                            Height =275
                            FontSize =10
                            Name ="Label1612"
                            Caption ="Pay-Off Inv Rcv"
                            LayoutCachedLeft =5490
                            LayoutCachedTop =3115
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =3390
                            RowStart =1
                            RowEnd =1
                            ColumnStart =11
                            ColumnEnd =11
                            BackThemeColorIndex =7
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6857
                    Top =2775
                    Width =1170
                    Height =270
                    TabIndex =23
                    ForeColor =-2147483600
                    Name ="PayOffRequested"
                    ControlSource ="PayOffRequested"
                    StatusBarText ="date invoice was requested from Accounting"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6857
                    LayoutCachedTop =2775
                    LayoutCachedWidth =8027
                    LayoutCachedHeight =3045
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =7
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =5490
                            Top =2775
                            Width =1290
                            Height =275
                            FontSize =10
                            Name ="Label1614"
                            Caption ="Pay-Off Inv Req"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =5490
                            LayoutCachedTop =2775
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =3050
                            ColumnStart =11
                            ColumnEnd =11
                            BackThemeColorIndex =7
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =8160
                    Top =1455
                    Width =968
                    Height =270
                    FontSize =10
                    Name ="Label1615"
                    Caption ="Deposit"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8160
                    LayoutCachedTop =1455
                    LayoutCachedWidth =9128
                    LayoutCachedHeight =1725
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =6
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8145
                    Top =60
                    Width =2340
                    Height =270
                    FontSize =10
                    TabIndex =38
                    Name ="UNSaleUnit"
                    Caption ="UN-Sell This Unit"
                    OnClick ="[Event Procedure]"
                    BottomPadding =165

                    LayoutCachedLeft =8145
                    LayoutCachedTop =60
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =330
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =2520
                    Top =795
                    Width =360
                    Height =1260
                    FontSize =10
                    BackColor =26316
                    ForeColor =16777215
                    Name ="Label1616"
                    Caption ="└ Buy Out ┘"
                    LeftPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =795
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =2055
                    RowStart =4
                    RowEnd =4
                    GridlineShade =65.0
                End
            End
        End
    End
End
CodeBehindForm
' See "Units_RMKT_Sub.cls"
