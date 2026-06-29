Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20160
    DatasheetFontHeight =11
    ItemSuffix =511
    Left =3720
    Top =7665
    Right =24885
    Bottom =20220
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0eac0befb57fe540
    End
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
            Height =1364
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =900
                    Width =20160
                    Height =464
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =900
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1364
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =95
                    Top =600
                    Width =20160
                    Height =300
                    BackColor =9211020
                    BorderColor =1643706
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =900
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =1008
                    Top =984
                    Width =1188
                    Height =300
                    Name ="Label1"
                    Caption ="Unit Number"
                    GroupTable =1
                    LayoutCachedLeft =1008
                    LayoutCachedTop =984
                    LayoutCachedWidth =2196
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =120
                    Top =984
                    Width =825
                    Height =300
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =10040879
                    Name ="Label20"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =984
                    LayoutCachedWidth =945
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =60
                    Top =984
                    Width =0
                    Height =300
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    LayoutCachedLeft =60
                    LayoutCachedTop =984
                    LayoutCachedWidth =60
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5940
                    Top =984
                    Width =1152
                    Height =300
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =5
                    Name ="BuyOutDateCbo"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =5940
                    LayoutCachedTop =984
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11400
                    Top =984
                    Width =1188
                    Height =300
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =8
                    Name ="IFinDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =11400
                    LayoutCachedTop =984
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15696
                    Top =984
                    Width =1152
                    Height =300
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =11
                    Name ="IFPaidDateCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =15696
                    LayoutCachedTop =984
                    LayoutCachedWidth =16848
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7164
                    Top =984
                    Width =1356
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =6
                    Name ="BuyOutPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =7164
                    LayoutCachedTop =984
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14412
                    Top =984
                    Width =1212
                    Height =300
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =10
                    Name ="IFAmountCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =14412
                    LayoutCachedTop =984
                    LayoutCachedWidth =15624
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =9360
                    Top =984
                    Width =1968
                    Height =300
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =7
                    Name ="IFBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName, Banks.BankName FROM Banks;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    Tag ="Medit"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =9360
                    LayoutCachedTop =984
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =8592
                            Top =984
                            Width =696
                            Height =300
                            FontSize =10
                            Name ="Label6"
                            Caption ="Status"
                            GroupTable =1
                            LayoutCachedLeft =8592
                            LayoutCachedTop =984
                            LayoutCachedWidth =9288
                            LayoutCachedHeight =1284
                            RowStart =1
                            RowEnd =1
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =9360
                    Top =600
                    Width =1968
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label46"
                    Caption =" Interim  Bank Name"
                    GroupTable =1
                    LayoutCachedLeft =9360
                    LayoutCachedTop =600
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =915
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =60
                    Top =600
                    Width =0
                    Height =315
                    Name ="EmptyCell68"
                    GroupTable =1
                    LayoutCachedLeft =60
                    LayoutCachedTop =600
                    LayoutCachedWidth =60
                    LayoutCachedHeight =915
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =120
                    Top =600
                    Width =825
                    Height =315
                    Name ="EmptyCell69"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =945
                    LayoutCachedHeight =915
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1008
                    Top =600
                    Width =1188
                    Height =315
                    Name ="EmptyCell70"
                    GroupTable =1
                    LayoutCachedLeft =1008
                    LayoutCachedTop =600
                    LayoutCachedWidth =2196
                    LayoutCachedHeight =915
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2256
                    Top =600
                    Width =2100
                    Height =315
                    Name ="EmptyCell71"
                    GroupTable =1
                    LayoutCachedLeft =2256
                    LayoutCachedTop =600
                    LayoutCachedWidth =4356
                    LayoutCachedHeight =915
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8592
                    Top =600
                    Width =696
                    Height =315
                    BackColor =1643706
                    Name ="EmptyCell75"
                    GroupTable =1
                    LayoutCachedLeft =8592
                    LayoutCachedTop =600
                    LayoutCachedWidth =9288
                    LayoutCachedHeight =915
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =7164
                    Top =600
                    Width =1356
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Buy Out Value"
                    GroupTable =1
                    LayoutCachedLeft =7164
                    LayoutCachedTop =600
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =915
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =5940
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="BuyOut Date"
                    GroupTable =1
                    LayoutCachedLeft =5940
                    LayoutCachedTop =600
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =915
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =15696
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label41"
                    Caption ="Interim Paid"
                    GroupTable =1
                    LayoutCachedLeft =15696
                    LayoutCachedTop =600
                    LayoutCachedWidth =16848
                    LayoutCachedHeight =915
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =14412
                    Top =600
                    Width =1212
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption =" Interim Amt"
                    GroupTable =1
                    LayoutCachedLeft =14412
                    LayoutCachedTop =600
                    LayoutCachedWidth =15624
                    LayoutCachedHeight =915
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =11400
                    Top =600
                    Width =1188
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption =" Interim Date"
                    GroupTable =1
                    LayoutCachedLeft =11400
                    LayoutCachedTop =600
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =915
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =12660
                    Top =600
                    Width =1680
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Loan Number"
                    GroupTable =1
                    LayoutCachedLeft =12660
                    LayoutCachedTop =600
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =915
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12660
                    Top =984
                    Width =1680
                    Height =300
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =9
                    Name ="IFAccountCBO"
                    GroupTable =1

                    LayoutCachedLeft =12660
                    LayoutCachedTop =984
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4428
                    Top =984
                    Height =300
                    FontSize =9
                    TabIndex =4
                    Name ="SellerInvCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =4428
                    LayoutCachedTop =984
                    LayoutCachedWidth =5868
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =2256
                            Top =984
                            Width =2100
                            Height =300
                            Name ="Label2"
                            Caption ="VIN"
                            GroupTable =1
                            LayoutCachedLeft =2256
                            LayoutCachedTop =984
                            LayoutCachedWidth =4356
                            LayoutCachedHeight =1284
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4428
                    Top =600
                    Width =1440
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Seller Invoice"
                    GroupTable =1
                    LayoutCachedLeft =4428
                    LayoutCachedTop =600
                    LayoutCachedWidth =5868
                    LayoutCachedHeight =915
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =4860
                    Top =360
                    Width =15264
                    Height =262
                    FontSize =9
                    FontWeight =700
                    Name ="Label417"
                    Caption =" ┌───────────── Buy Out ──────────────┐                                         "
                        "┌──────────────────────── Interim Remarketing Loan ───────────────────────┐     "
                        "             ┌── Bank Invoice PayOff ──┐   "
                    LayoutCachedLeft =4860
                    LayoutCachedTop =360
                    LayoutCachedWidth =20124
                    LayoutCachedHeight =622
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17112
                    Top =984
                    Width =1152
                    Height =300
                    FontSize =9
                    TabIndex =12
                    Name ="PayOffRequestedCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =17112
                    LayoutCachedTop =984
                    LayoutCachedWidth =18264
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =16908
                            Top =984
                            Width =132
                            Height =300
                            FontSize =10
                            Name ="Label495"
                            Caption =" "
                            GroupTable =1
                            LayoutCachedLeft =16908
                            LayoutCachedTop =984
                            LayoutCachedWidth =17040
                            LayoutCachedHeight =1284
                            RowStart =1
                            RowEnd =1
                            ColumnStart =13
                            ColumnEnd =13
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18336
                    Top =984
                    Width =1152
                    Height =300
                    FontSize =9
                    TabIndex =13
                    Name ="PayOffReceivedCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =18336
                    LayoutCachedTop =984
                    LayoutCachedWidth =19488
                    LayoutCachedHeight =1284
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16980
                    Top =60
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="PayOffRequestedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =5
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16980
                    LayoutCachedTop =60
                    LayoutCachedWidth =18135
                    LayoutCachedHeight =300
                    LayoutGroup =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =7883872
                    BorderColor =6108695
                    HoverColor =7883872
                    PressedColor =7883872
                    GroupTable =5
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =18192
                    Top =60
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="PayOffReceivedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18192
                    LayoutCachedTop =60
                    LayoutCachedWidth =19344
                    LayoutCachedHeight =300
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =9732411
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26265
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =17112
                    Top =600
                    Width =1152
                    Height =315
                    ForeColor =16777215
                    Name ="Label491"
                    Caption ="Rquested"
                    GroupTable =1
                    LayoutCachedLeft =17112
                    LayoutCachedTop =600
                    LayoutCachedWidth =18264
                    LayoutCachedHeight =915
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =18336
                    Top =600
                    Width =1152
                    Height =315
                    ForeColor =16777215
                    Name ="Label493"
                    Caption ="Received"
                    GroupTable =1
                    LayoutCachedLeft =18336
                    LayoutCachedTop =600
                    LayoutCachedWidth =19488
                    LayoutCachedHeight =915
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16908
                    Top =600
                    Width =132
                    Height =315
                    BackColor =1643706
                    Name ="EmptyCell497"
                    GroupTable =1
                    LayoutCachedLeft =16908
                    LayoutCachedTop =600
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =915
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =13980
                    Top =60
                    Width =2835
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="MakeBankPayOffEmail"
                    Caption =" Make Bank Invoice Request Email"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =10
                    RightPadding =40
                    BottomPadding =39
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =13980
                    LayoutCachedTop =60
                    LayoutCachedWidth =16815
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderWidth =1
                    BorderColor =62207
                    HoverColor =5676533
                    PressedColor =26265
                    GroupTable =10
                    QuickStyle =31
                    QuickStyleMask =-245
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =19500
                    Top =972
                    Width =405
                    Height =300
                    FontSize =9
                    Name ="Label510"
                    Caption ="Title"
                    GroupTable =12
                    LayoutCachedLeft =19500
                    LayoutCachedTop =972
                    LayoutCachedWidth =19905
                    LayoutCachedHeight =1272
                    LayoutGroup =6
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =60
                    Top =36
                    Width =0
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =36
                    LayoutCachedWidth =60
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1008
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =3
                    Name ="UnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =1

                    LayoutCachedLeft =1008
                    LayoutCachedTop =36
                    LayoutCachedWidth =2196
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2256
                    Top =36
                    Width =2100
                    Height =252
                    FontSize =10
                    TabIndex =4
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =2256
                    LayoutCachedTop =36
                    LayoutCachedWidth =4356
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5940
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =6
                    BackColor =13237503
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    ValidationText ="BO Date is required"
                    GroupTable =1

                    LayoutCachedLeft =5940
                    LayoutCachedTop =36
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7164
                    Top =36
                    Width =1356
                    Height =252
                    FontSize =10
                    TabIndex =7
                    BackColor =13237503
                    Name ="BuyOutCost"
                    ControlSource ="BuyOutCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =7164
                    LayoutCachedTop =36
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12660
                    Top =36
                    Width =1680
                    Height =252
                    FontSize =10
                    TabIndex =11
                    BackColor =16381933
                    Name ="IFLoanNumber"
                    ControlSource ="IFLoanNumber"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =12660
                    LayoutCachedTop =36
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8592
                    Top =36
                    Width =696
                    Height =252
                    FontSize =10
                    TabIndex =8
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1

                    LayoutCachedLeft =8592
                    LayoutCachedTop =36
                    LayoutCachedWidth =9288
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =36
                    Width =825
                    Height =252
                    FontSize =10
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =10040879
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =36
                    LayoutCachedWidth =945
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11400
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =10
                    BackColor =16317438
                    Name ="IFDate"
                    ControlSource ="IFDate"
                    GroupTable =1

                    LayoutCachedLeft =11400
                    LayoutCachedTop =36
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15696
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =13
                    BackColor =16317438
                    Name ="IFPaidDate"
                    ControlSource ="IFPaidDate"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =15696
                    LayoutCachedTop =36
                    LayoutCachedWidth =16848
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =9360
                    Top =36
                    Width =1968
                    Height =252
                    FontSize =10
                    TabIndex =9
                    BackColor =16317438
                    Name ="UnitIFBank"
                    ControlSource ="IFBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName, Banks.BankName FROM Banks;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =9360
                    LayoutCachedTop =36
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =20160
                    Height =59
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =315
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4428
                    Top =36
                    Height =252
                    FontSize =10
                    TabIndex =5
                    Name ="SellerInvoice"
                    ControlSource ="SellerInvoice"
                    GroupTable =1

                    LayoutCachedLeft =4428
                    LayoutCachedTop =36
                    LayoutCachedWidth =5868
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14412
                    Top =36
                    Width =1212
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BackColor =16317438
                    Name ="IFAmount"
                    ControlSource ="IFAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =14412
                    LayoutCachedTop =36
                    LayoutCachedWidth =15624
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17112
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =14
                    BackColor =13611711
                    Name ="PayOffRequested"
                    ControlSource ="PayOffRequested"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =17112
                    LayoutCachedTop =36
                    LayoutCachedWidth =18264
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18336
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =15
                    Name ="PayOffReceived"
                    ControlSource ="PayOffReceived"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =18336
                    LayoutCachedTop =36
                    LayoutCachedWidth =19488
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16908
                    Top =36
                    Width =132
                    Height =252
                    BackColor =1643706
                    Name ="EmptyCell499"
                    GroupTable =1
                    LayoutCachedLeft =16908
                    LayoutCachedTop =36
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =288
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19500
                    Top =36
                    Width =465
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="LastDirectionx"
                    ControlSource ="Last"
                    Format ="Short Date"
                    ControlTipText ="You may change this to the date needed"
                    GroupTable =11

                    LayoutCachedLeft =19500
                    LayoutCachedTop =36
                    LayoutCachedWidth =19965
                    LayoutCachedHeight =312
                    LayoutGroup =5
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
            End
        End
        Begin FormFooter
            Height =420
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4530
                    Width =29
                    Height =270
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =2
                    Name ="SchNo"
                    ControlSource ="SchNo"

                    LayoutCachedLeft =4530
                    LayoutCachedWidth =4559
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =4410
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =4410
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4260
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =4260
                    LayoutCachedWidth =4289
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14328
                    Height =255
                    FontSize =10
                    TabIndex =3
                    BackColor =16317438
                    Name ="Text225"
                    ControlSource ="=Sum([IFAmount])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =14328
                    LayoutCachedWidth =15768
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6285
                    Width =750
                    Height =255
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Totals:"
                    LayoutCachedLeft =6285
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7098
                    Height =255
                    FontSize =10
                    TabIndex =4
                    BackColor =13237503
                    Name ="Text14"
                    ControlSource ="=Sum([BuyOutCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7098
                    LayoutCachedWidth =8538
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
            End
        End
    End
End
CodeBehindForm
' See "ReMarketingMassMod_sub_BuyOutView.cls"
