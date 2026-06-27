Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =1
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =284
    Left =5460
    Top =17265
    Right =18345
    Bottom =21045
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6dcbcf5ff29fe540
    End
    RecordSource ="SELECT * FROM RemarketExpense WHERE (((RemarketExpense.[RMKTExpenseId])=0));"
    BeforeUpdate ="[Event Procedure]"
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
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =570
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =360
                    Width =12960
                    Height =210
                    BorderColor =16777215
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =360
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =570
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Width =7590
                    Height =420
                    FontSize =16
                    Name ="Label1053"
                    Caption ="Remarketing Expenses for this Unit"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
            End
        End
        Begin Section
            Height =3600
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =7545
                    Top =765
                    Width =2070
                    FontSize =10
                    TabIndex =5
                    Name ="VendorRef"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName FROM Vendors WHER"
                        "E (((Vendors.VendorObsolete)=False)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;1440;4320"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =765
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =1005
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =765
                            Width =1425
                            Height =240
                            FontSize =9
                            Name ="Label2"
                            Caption ="Vendor"
                            LayoutCachedLeft =6060
                            LayoutCachedTop =765
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =1005
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7545
                    Top =1680
                    Width =2070
                    FontSize =10
                    TabIndex =8
                    Name ="InvoiceNumber"
                    ControlSource ="InvoiceNumber"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =1920
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =1680
                            Width =1425
                            Height =240
                            FontSize =9
                            Name ="Label3"
                            Caption ="Vendor Inv No."
                            LayoutCachedLeft =6060
                            LayoutCachedTop =1680
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =1920
                            RowStart =4
                            RowEnd =4
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7545
                    Top =1980
                    Width =2070
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="InvoiceDate"
                    ControlSource ="InvoiceDate"
                    Format ="Short Date"
                    BottomPadding =45
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Condition ="DatePart(\"yyyy\",Date())<>DatePart(\"yyyy\",[InvoiceDate])"
                            Action ="MsgBox"
                            Argument ="Invoice Date not in current year. Is this correct?"
                            Argument ="-1"
                            Argument ="2"
                            Argument ="Check Invoice Date"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"InvoiceDate\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition>"
                                "DatePart(\"yyyy\",Date())&lt;&g"
                        End
                        Begin
                            Comment ="_AXL:t;DatePart(\"yyyy\",[InvoiceDate])</Condition><Statements><Action Name=\"Me"
                                "ssageBox\"><Argument Name=\"Message\">Invoice Date not in current year. Is this "
                                "correct?</Argument><Argument Name=\"Type\">Warning?</Argument><Argument Name=\"T"
                                "itle\">Check Invoice Da"
                        End
                        Begin
                            Comment ="_AXL:te</Argument></Action></Statements></If></ConditionalBlock></Statements></U"
                                "serInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7545
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =2250
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =1980
                            Width =1425
                            Height =270
                            FontSize =9
                            Name ="Label4"
                            Caption ="Vendor Inv Date:"
                            BottomPadding =45
                            LayoutCachedLeft =6060
                            LayoutCachedTop =1980
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =2250
                            RowStart =5
                            RowEnd =5
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7560
                    Top =2340
                    Width =5280
                    Height =720
                    FontSize =10
                    TabIndex =13
                    Name ="Comment"
                    ControlSource ="Comment"
                    BottomPadding =45

                    LayoutCachedLeft =7560
                    LayoutCachedTop =2340
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =3060
                    RowStart =1
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3480
                    Top =3060
                    Width =1320
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="RMKTExpenseId"
                    ControlSource ="RMKTExpenseId"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =3480
                    LayoutCachedTop =3060
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =3300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =3060
                    Width =780
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    Name ="AssetId"
                    ControlSource ="AssetId"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =60
                    LayoutCachedTop =3060
                    LayoutCachedWidth =840
                    LayoutCachedHeight =3330
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =7545
                    Top =420
                    Width =2070
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="ExpenseType"
                    ControlSource ="ExpenseTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketExpenseType.RemarketExpenseTypeId, RemarketExpenseType.RemarketEx"
                        "pense FROM RemarketExpenseType ORDER BY RemarketExpenseType.RemarketExpense;"
                    ColumnWidths ="0;1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Expense type is required"
                    ControlTipText ="Double click to add expsenst type"
                    BottomPadding =45
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
                                "nterfaceMacro For=\"ExpenseType\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argume"
                                "nt Name=\"FormName\">AddRmktExpFo"
                        End
                        Begin
                            Comment ="_AXL:rm</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7545
                    LayoutCachedTop =420
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =420
                            Width =1425
                            Height =270
                            FontSize =9
                            Name ="Label6"
                            Caption ="Expense Type"
                            ControlTipText ="Double click to add expsenst type"
                            BottomPadding =45
                            LayoutCachedLeft =6060
                            LayoutCachedTop =420
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =690
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =60
                    Top =315
                    Width =5814
                    Height =2520
                    FontSize =9
                    TabIndex =3
                    Name ="ExpenseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT R.RMKTExpenseId, V.VendorSNm AS Vendor, T.RemarketExpense AS Type, CCur(R"
                        ".Expense) AS Expense, L.ReimburseStatus AS RStatus, Format(ReimburseDate,\"mm/dd"
                        "/yy\") AS RDate FROM ((RemarketExpense AS R INNER JOIN RemarketExpenseType AS T "
                        "ON R.ExpenseTypeId = T.RemarketExpenseTypeId) INNER JOIN LookUp_ReimburseStatus "
                        "AS L ON R.ReimburseStatusID = L.ReimburseStatusID) LEFT JOIN Vendors AS V ON R.V"
                        "endorId = V.VendorID WHERE (((R.[AssetId])=0)) ORDER BY T.RemarketExpense;"
                    ColumnWidths ="0;1296;936;893;1584;864"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =315
                    LayoutCachedWidth =5874
                    LayoutCachedHeight =2835
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7545
                    Top =1065
                    Width =2070
                    FontSize =10
                    TabIndex =6
                    Name ="ExpenseDate"
                    ControlSource ="ExpenseDate"
                    Format ="Short Date"
                    BottomPadding =45
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
                                "nterfaceMacro For=\"ExpenseDate\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argume"
                                "nt Name=\"FormName\">AddRmktExpFo"
                        End
                        Begin
                            Comment ="_AXL:rm</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7545
                    LayoutCachedTop =1065
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =1305
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =1065
                            Width =1425
                            Height =240
                            FontSize =9
                            Name ="Label108"
                            Caption ="Expense Date"
                            BottomPadding =45
                            LayoutCachedLeft =6060
                            LayoutCachedTop =1065
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =1305
                            RowStart =2
                            RowEnd =2
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6060
                    Top =2325
                    Width =1425
                    Height =270
                    FontSize =9
                    Name ="Label152"
                    Caption ="Comments"
                    BottomPadding =45
                    LayoutCachedLeft =6060
                    LayoutCachedTop =2325
                    LayoutCachedWidth =7485
                    LayoutCachedHeight =2595
                    RowStart =6
                    RowEnd =6
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2292
                    Top =2880
                    Width =879
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    ForeColor =10040879
                    Name ="ExpenseSum"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =2292
                    LayoutCachedTop =2880
                    LayoutCachedWidth =3171
                    LayoutCachedHeight =3120
                    DatasheetCaption ="Final Cost"
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1140
                            Top =2880
                            Width =1155
                            Height =240
                            FontSize =8
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label573"
                            Caption ="Total Expense"
                            LayoutCachedLeft =1140
                            LayoutCachedTop =2880
                            LayoutCachedWidth =2295
                            LayoutCachedHeight =3120
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =2
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7545
                    Top =1380
                    Width =2070
                    FontSize =10
                    TabIndex =7
                    Name ="Expense"
                    ControlSource ="Expense"
                    Format ="$#,##0.00;($#,##0.00)"
                    DefaultValue ="0"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =1380
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =1380
                            Width =1425
                            Height =240
                            FontSize =9
                            Name ="Label179"
                            Caption ="Expense Amt"
                            LayoutCachedLeft =6060
                            LayoutCachedTop =1380
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =1620
                            RowStart =3
                            RowEnd =3
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3900
                    Top =60
                    Width =2685
                    Height =210
                    FontSize =10
                    FontWeight =700
                    Name ="EnterMsg"
                    Caption ="▼Enter Expense Data HERE▼"
                    GroupTable =2
                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =6585
                    LayoutCachedHeight =270
                    LayoutGroup =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11100
                    Top =1380
                    Width =1788
                    FontSize =10
                    TabIndex =10
                    Name ="ReimburseStatusID"
                    ControlSource ="ReimburseStatusID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ReimburseStatus.ReimburseStatusID, LookUp_ReimburseStatus.Reimburs"
                        "eStatus FROM LookUp_ReimburseStatus;"
                    ColumnWidths ="0;1440"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =1380
                    LayoutCachedWidth =12888
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =7
                    BackTint =20.0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =9705
                    Top =1380
                    Width =1320
                    Height =240
                    FontSize =9
                    Name ="Label310"
                    Caption ="Reimbursed"
                    LayoutCachedLeft =9705
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11025
                    LayoutCachedHeight =1620
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =9705
                    Top =1680
                    Width =1320
                    Height =240
                    FontSize =9
                    Name ="Label43"
                    Caption ="R Event Date"
                    LayoutCachedLeft =9705
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11025
                    LayoutCachedHeight =1920
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11100
                    Top =1680
                    Width =1788
                    FontSize =10
                    TabIndex =11
                    Name ="ReimburseDate"
                    ControlSource ="ReimburseDate"
                    Format ="Short Date"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12888
                    LayoutCachedHeight =1920
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =7
                    BackTint =20.0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11220
                    Width =1560
                    Height =240
                    FontSize =10
                    Name ="DeleteMe"
                    Caption ="Delete Expense"
                    OnClick ="[Event Procedure]"
                    BottomPadding =45

                    LayoutCachedLeft =11220
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9120
                    Top =15
                    Width =1560
                    Height =240
                    FontSize =10
                    TabIndex =2
                    Name ="AddNew"
                    Caption ="Add An Expense"
                    OnClick ="[Event Procedure]"
                    BottomPadding =45

                    LayoutCachedLeft =9120
                    LayoutCachedTop =15
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =50.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7380
                    Top =15
                    Width =1560
                    Height =240
                    FontSize =10
                    TabIndex =1
                    Name ="SaveMe"
                    Caption ="Update / Save "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =15
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =2
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11355
                    Top =3060
                    Width =1500
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    ForeColor =8355711
                    Name ="DateExpenseAdded"
                    ControlSource ="DateExpenseAdded"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =11355
                    LayoutCachedTop =3060
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =3330
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =10740
                            Top =3060
                            Width =600
                            Height =270
                            FontSize =9
                            Name ="Label283"
                            Caption ="Added"
                            LayoutCachedLeft =10740
                            LayoutCachedTop =3060
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =3330
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =9705
                    Top =1980
                    Width =1320
                    Height =270
                    FontSize =9
                    Name ="Label226"
                    Caption ="R Invoice Num"
                    BottomPadding =45
                    LayoutCachedLeft =9705
                    LayoutCachedTop =1980
                    LayoutCachedWidth =11025
                    LayoutCachedHeight =2250
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11100
                    Top =1980
                    Width =1788
                    Height =270
                    FontSize =10
                    TabIndex =12
                    Name ="RInvoiceNum"
                    ControlSource ="RInvoiceNum"
                    BottomPadding =45

                    LayoutCachedLeft =11100
                    LayoutCachedTop =1980
                    LayoutCachedWidth =12888
                    LayoutCachedHeight =2250
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =7
                    BackTint =20.0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "RmktExpense.cls"
