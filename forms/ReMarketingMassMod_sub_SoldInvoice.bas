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
    Width =19433
    DatasheetFontHeight =11
    ItemSuffix =576
    Left =6720
    Top =6270
    Right =22035
    Bottom =18825
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xaf291a07b67fe540
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
            Height =1296
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =900
                    Width =18360
                    Height =396
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =900
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1296
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =600
                    Width =19410
                    Height =285
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =19410
                    LayoutCachedHeight =885
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =996
                    Top =960
                    Width =1188
                    Height =288
                    Name ="Label1"
                    Caption ="Unit Num"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =960
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
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
                    Left =96
                    Top =960
                    Width =828
                    Height =288
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =10040879
                    Name ="Label20"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =960
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =36
                    Top =960
                    Width =0
                    Height =288
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =960
                    LayoutCachedWidth =36
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6468
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="RmktSPRAdjSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6468
                    LayoutCachedTop =300
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =540
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6468
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =15
                    Name ="RmktSPRAdjCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =6468
                    LayoutCachedTop =960
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =18084
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="RmktBOSSentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18084
                    LayoutCachedTop =300
                    LayoutCachedWidth =19236
                    LayoutCachedHeight =540
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18084
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =25
                    BorderColor =9211020
                    Name ="RmktBOSSentCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =18084
                    LayoutCachedTop =960
                    LayoutCachedWidth =19236
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7692
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="SoldPriceSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7692
                    LayoutCachedTop =300
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =540
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7692
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =16
                    Name ="SoldPriceCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =7692
                    LayoutCachedTop =960
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11364
                    Top =300
                    Width =840
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="RmkSmSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11364
                    LayoutCachedTop =300
                    LayoutCachedWidth =12204
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =8
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =996
                    Top =300
                    Width =1188
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="send2Excel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =300
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =540
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =36
                    Top =600
                    Width =0
                    Height =285
                    Name ="EmptyCell68"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =600
                    LayoutCachedWidth =36
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =600
                    Width =828
                    Height =285
                    Name ="EmptyCell69"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =600
                    LayoutCachedWidth =924
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =996
                    Top =600
                    Width =1188
                    Height =285
                    Name ="EmptyCell70"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =600
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2244
                    Top =600
                    Width =2100
                    Height =285
                    Name ="EmptyCell71"
                    GroupTable =1
                    LayoutCachedLeft =2244
                    LayoutCachedTop =600
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7692
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Sold Price"
                    GroupTable =1
                    LayoutCachedLeft =7692
                    LayoutCachedTop =600
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =300
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =300
                    LayoutCachedWidth =36
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =300
                    Width =828
                    Name ="EmptyCell97"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =300
                    LayoutCachedWidth =924
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2244
                    Top =300
                    Width =2100
                    Name ="EmptyCell99"
                    GroupTable =1
                    LayoutCachedLeft =2244
                    LayoutCachedTop =300
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =540
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6468
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Adjusted"
                    GroupTable =1
                    LayoutCachedLeft =6468
                    LayoutCachedTop =600
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =18084
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label41"
                    Caption ="BOS Sent"
                    GroupTable =1
                    LayoutCachedLeft =18084
                    LayoutCachedTop =600
                    LayoutCachedWidth =19236
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =11364
                    Top =600
                    Width =840
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption ="Rep"
                    GroupTable =1
                    LayoutCachedLeft =11364
                    LayoutCachedTop =600
                    LayoutCachedWidth =12204
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10140
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Received"
                    GroupTable =1
                    LayoutCachedLeft =10140
                    LayoutCachedTop =600
                    LayoutCachedWidth =11292
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10140
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="RmktDepdtSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10140
                    LayoutCachedTop =300
                    LayoutCachedWidth =11292
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10140
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =18
                    Name ="RmktDepdtCBO"
                    Format ="Short Date"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =10140
                    LayoutCachedTop =960
                    LayoutCachedWidth =11292
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8916
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =17
                    Name ="RmktDepositCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =8916
                    LayoutCachedTop =960
                    LayoutCachedWidth =10068
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8916
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Amount"
                    GroupTable =1
                    LayoutCachedLeft =8916
                    LayoutCachedTop =600
                    LayoutCachedWidth =10068
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8916
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="RmktDepositSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8916
                    LayoutCachedTop =300
                    LayoutCachedWidth =10068
                    LayoutCachedHeight =540
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4416
                    Top =300
                    Width =1980
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="BuyerRefSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4416
                    LayoutCachedTop =300
                    LayoutCachedWidth =6396
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =15636
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="InvoiceSentDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15636
                    LayoutCachedTop =300
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =540
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =15636
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label379"
                    Caption ="Sent"
                    GroupTable =1
                    LayoutCachedLeft =15636
                    LayoutCachedTop =600
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4416
                    Top =600
                    Width =1980
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label381"
                    Caption ="Buyer"
                    GroupTable =1
                    LayoutCachedLeft =4416
                    LayoutCachedTop =600
                    LayoutCachedWidth =6396
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =14412
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label383"
                    Caption ="Requested"
                    GroupTable =1
                    LayoutCachedLeft =14412
                    LayoutCachedTop =600
                    LayoutCachedWidth =15564
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15636
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =23
                    BorderColor =9211020
                    Name ="InvoiceSentDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =15636
                    LayoutCachedTop =960
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6780
                    Top =60
                    Width =4245
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption =" ┌───── Price ─────┐                  ┌─────  Deposit ─────┐"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =11025
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =14760
                    Top =60
                    Width =4228
                    Height =262
                    FontSize =9
                    FontWeight =700
                    Name ="Label417"
                    Caption =" ┌──────────────── Buyers Order ───────────────┐"
                    LayoutCachedLeft =14760
                    LayoutCachedTop =60
                    LayoutCachedWidth =18988
                    LayoutCachedHeight =322
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =16860
                    Top =600
                    Width =1152
                    Height =285
                    BackColor =9732411
                    Name ="Label426"
                    Caption ="Paid"
                    GroupTable =1
                    LayoutCachedLeft =16860
                    LayoutCachedTop =600
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14412
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =22
                    BorderColor =9211020
                    Name ="InvoiceRequestDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =14412
                    LayoutCachedTop =960
                    LayoutCachedWidth =15564
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14412
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="InvoiceRequestDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14412
                    LayoutCachedTop =300
                    LayoutCachedWidth =15564
                    LayoutCachedHeight =540
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =7200
                    Left =4416
                    Top =960
                    Width =1980
                    Height =288
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =14
                    Name ="BuyerRefCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.clientID, Clients.ClientShNm, Clients.clientCompanyName, IIf([cli"
                        "enttype]=\"Remarketing\",\"R\",\"C\") AS Type FROM Clients WHERE (((Clients.Clie"
                        "ntShNm) Not Like \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.cli"
                        "entType)=\"guarantor\" Or (Clients.clientType)=\"remarketing\")) ORDER BY IIf([c"
                        "lienttype]=\"Remarketing\",\"r\",\"C\") DESC , Clients.ClientShNm, Clients.clien"
                        "tType DESC , Clients.clientshnm;"
                    ColumnWidths ="0;2160;4320;288"
                    StatusBarText ="Link to Buyer"
                    ControlTipText ="This is the Buyer, Double Click to view"
                    GroupTable =1

                    LayoutCachedLeft =4416
                    LayoutCachedTop =960
                    LayoutCachedWidth =6396
                    LayoutCachedHeight =1248
                    DisplayAsHyperlink =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =2244
                            Top =960
                            Width =2100
                            Height =288
                            Name ="Label2"
                            Caption ="VIN"
                            GroupTable =1
                            LayoutCachedLeft =2244
                            LayoutCachedTop =960
                            LayoutCachedWidth =4344
                            LayoutCachedHeight =1248
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16860
                    Top =960
                    Width =1152
                    Height =288
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =24
                    BorderColor =9211020
                    Name ="InvoicePaidDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =16860
                    LayoutCachedTop =960
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =16860
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="InvoicePaidDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16860
                    LayoutCachedTop =300
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =540
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11364
                    Top =960
                    Width =840
                    Height =288
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =19
                    Name ="RmkSmCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.OBSOLETE, Emps.Rmkt FROM Emps WHERE (((Emps.OBSOLETE)=N"
                        "o) AND ((Emps.Rmkt)=Yes)) ORDER BY Emps.EmpInt;"
                    StatusBarText ="initials of Remarketing Salesperson"
                    GroupTable =1

                    LayoutCachedLeft =11364
                    LayoutCachedTop =960
                    LayoutCachedWidth =12204
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =12276
                    Top =960
                    Width =852
                    Height =288
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =20
                    Name ="UnitStatusCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="432;1440"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =12276
                    LayoutCachedTop =960
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12276
                    Top =600
                    Width =852
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label548"
                    Caption ="Status"
                    GroupTable =1
                    LayoutCachedLeft =12276
                    LayoutCachedTop =600
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12276
                    Top =300
                    Width =852
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="UnitStatusSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12276
                    LayoutCachedTop =300
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =13200
                    Top =960
                    Width =1140
                    Height =288
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =21
                    Name ="WholesaleCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_RMKTDisposition.DispositionID, Lookup_RMKTDisposition.DispositionT"
                        "ype FROM Lookup_RMKTDisposition;"
                    ColumnWidths ="0;1440"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =13200
                    LayoutCachedTop =960
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =1248
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =13200
                    Top =600
                    Width =1140
                    Height =285
                    ForeColor =16777215
                    Name ="Label570"
                    Caption ="Disposition"
                    GroupTable =1
                    LayoutCachedLeft =13200
                    LayoutCachedTop =600
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13200
                    Top =300
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="WholesaleSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13200
                    LayoutCachedTop =300
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =540
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
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
                    Left =36
                    Top =36
                    Width =0
                    Height =252
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =36
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =36
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
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
                    Left =2244
                    Top =36
                    Width =2100
                    Height =252
                    FontSize =10
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =2244
                    LayoutCachedTop =36
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6468
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =5
                    BackColor =15070975
                    Name ="RmktSPRAdj"
                    ControlSource ="RmktSPRAdj"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Adjustment made to the unit's sale price"
                    GroupTable =1

                    LayoutCachedLeft =6468
                    LayoutCachedTop =36
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7692
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =6
                    BackColor =15070975
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Sold Price of the Unit"
                    GroupTable =1

                    LayoutCachedLeft =7692
                    LayoutCachedTop =36
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10140
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =8
                    BackColor =15070975
                    Name ="MiParked"
                    ControlSource ="RmktDepdt"
                    Format ="Short Date"
                    DefaultValue ="0"
                    ControlTipText ="Date the deposit was made"
                    GroupTable =1

                    LayoutCachedLeft =10140
                    LayoutCachedTop =36
                    LayoutCachedWidth =11292
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =96
                    Top =36
                    Width =828
                    Height =252
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =10040879
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =96
                    LayoutCachedTop =36
                    LayoutCachedWidth =924
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =315
                    Width =18360
                    Height =59
                    FontSize =10
                    TabIndex =16
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
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8916
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =7
                    BackColor =15070975
                    Name ="RmktDeposit"
                    ControlSource ="RmktDeposit"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Buyer's Deposit towrds the sale of the unit"
                    GroupTable =1

                    LayoutCachedLeft =8916
                    LayoutCachedTop =36
                    LayoutCachedWidth =10068
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15636
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =13
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date the Buyer was Invoiced for the unit"
                    GroupTable =1

                    LayoutCachedLeft =15636
                    LayoutCachedTop =36
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18084
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =15
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="RmktBOSSent"
                    ControlSource ="RmktBOSSent"
                    Format ="Short Date"
                    ControlTipText ="Date the Buyer was sent the Bill Of Sale for the unit"
                    GroupTable =1

                    LayoutCachedLeft =18084
                    LayoutCachedTop =36
                    LayoutCachedWidth =19236
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14412
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="InvoiceRequestDate"
                    ControlSource ="InvoiceRequestDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date an sold Invoice was requested"
                    GroupTable =1

                    LayoutCachedLeft =14412
                    LayoutCachedTop =36
                    LayoutCachedWidth =15564
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11364
                    Top =36
                    Width =840
                    Height =252
                    FontSize =10
                    TabIndex =9
                    Name ="RmkSm"
                    ControlSource ="RmkSm"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.OBSOLETE, Emps.Rmkt FROM Emps WHERE (((Emps.OBSOLETE)=N"
                        "o) AND ((Emps.Rmkt)=Yes)) ORDER BY Emps.EmpInt;"
                    StatusBarText ="initials of Remarketing Salesperson"
                    GroupTable =1

                    LayoutCachedLeft =11364
                    LayoutCachedTop =36
                    LayoutCachedWidth =12204
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =7200
                    Left =4416
                    Top =36
                    Width =1980
                    Height =252
                    FontSize =10
                    TabIndex =4
                    Name ="BuyerRef"
                    ControlSource ="BuyerRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.clientID, Clients.ClientShNm, Clients.clientCompanyName, IIf([cli"
                        "enttype]=\"Remarketing\",\"R\",\"C\") AS Type FROM Clients WHERE (((Clients.Clie"
                        "ntShNm) Not Like \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.cli"
                        "entType)=\"guarantor\" Or (Clients.clientType)=\"remarketing\")) ORDER BY IIf([c"
                        "lienttype]=\"Remarketing\",\"r\",\"C\") DESC , Clients.ClientShNm, Clients.clien"
                        "tType DESC , Clients.clientshnm;"
                    ColumnWidths ="0;2160;4320;288"
                    StatusBarText ="Link to Buyer"
                    ControlTipText ="This is the Buyer, Double Click to view"
                    GroupTable =1

                    LayoutCachedLeft =4416
                    LayoutCachedTop =36
                    LayoutCachedWidth =6396
                    LayoutCachedHeight =288
                    DisplayAsHyperlink =1
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16860
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =14
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="InvoicePaidDate"
                    ControlSource ="InvoicePaidDate"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date the Invoice was Paid"
                    GroupTable =1

                    LayoutCachedLeft =16860
                    LayoutCachedTop =36
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =12276
                    Top =36
                    Width =852
                    Height =252
                    FontSize =10
                    TabIndex =10
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="432;1440"
                    StatusBarText ="Unit Status"
                    GroupTable =1

                    LayoutCachedLeft =12276
                    LayoutCachedTop =36
                    LayoutCachedWidth =13128
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =13200
                    Top =36
                    Width =1140
                    Height =252
                    FontSize =10
                    TabIndex =11
                    Name ="Wholesale"
                    ControlSource ="Wholesale"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_RMKTDisposition.DispositionID, Lookup_RMKTDisposition.DispositionT"
                        "ype FROM Lookup_RMKTDisposition;"
                    ColumnWidths ="0;1440"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =13200
                    LayoutCachedTop =36
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9810
                    Width =29
                    Height =270
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =2
                    Name ="SchNo"
                    ControlSource ="SchNo"

                    LayoutCachedLeft =9810
                    LayoutCachedWidth =9839
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =9690
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =9690
                    LayoutCachedWidth =9719
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9540
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =9540
                    LayoutCachedWidth =9569
                    LayoutCachedHeight =270
                End
            End
        End
    End
End
CodeBehindForm
' See "ReMarketingMassMod_sub_SoldInvoice.cls"
