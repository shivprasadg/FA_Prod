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
    Width =19536
    DatasheetFontHeight =11
    ItemSuffix =728
    Left =6720
    Top =6270
    Right =22035
    Bottom =18825
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x91d5ad7eb57fe540
    End
    RecordSource ="RemarketingMassMod1"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            Height =1302
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4725
                    Top =60
                    Width =2175
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label88"
                    Caption ="  ─ Buy Out  Book Value ─┐        "
                    LayoutCachedLeft =4725
                    LayoutCachedTop =60
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =285
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =900
                    Width =19425
                    Height =362
                    Name ="Box48"
                    LayoutCachedTop =900
                    LayoutCachedWidth =19425
                    LayoutCachedHeight =1262
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =600
                    Width =19425
                    Height =285
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =19425
                    LayoutCachedHeight =885
                    BackThemeColorIndex =1
                    BackShade =65.0
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
                    Height =312
                    Name ="Label1"
                    Caption ="Unit Num"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =960
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =1272
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
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =2244
                    Top =960
                    Width =2100
                    Height =312
                    Name ="Label2"
                    Caption ="VIN"
                    GroupTable =1
                    LayoutCachedLeft =2244
                    LayoutCachedTop =960
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
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
                    Height =312
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
                    LayoutCachedHeight =1272
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
                    Height =312
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =960
                    LayoutCachedWidth =36
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12348
                    Top =300
                    Width =1236
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="WhSValueSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12348
                    LayoutCachedTop =300
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =6108695
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12348
                    Top =960
                    Width =1236
                    Height =312
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =12
                    BackColor =16317438
                    Name ="WhSValueCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =12348
                    LayoutCachedTop =960
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =11004
                            Top =960
                            Width =1272
                            Height =312
                            BackColor =1643706
                            Name ="Label579"
                            Caption ="GainLoss"
                            GroupTable =1
                            LayoutCachedLeft =11004
                            LayoutCachedTop =960
                            LayoutCachedWidth =12276
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =18600
                    Top =300
                    Width =900
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="FABrkCommSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18600
                    LayoutCachedTop =300
                    LayoutCachedWidth =19500
                    LayoutCachedHeight =540
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10040879
                    BorderColor =6108695
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedColor =10040879
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18600
                    Top =960
                    Width =900
                    Height =312
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =16
                    Name ="FABrkCommCbo"
                    Format ="Percent"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =18600
                    LayoutCachedTop =960
                    LayoutCachedWidth =19500
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5700
                    Top =300
                    Width =1224
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="CopyBOSubmit"
                    Caption ="▼ Copy BO"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Copy the Buy-Out to Book Value, will not overwrite exisiting values."
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5700
                    LayoutCachedTop =300
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =540
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =7883872
                    BorderColor =6108695
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedColor =1643706
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5700
                    Top =960
                    Width =1224
                    Height =312
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =10
                    Name ="AdjBookValueCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =5700
                    LayoutCachedTop =960
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =4404
                            Top =960
                            Width =1224
                            Height =312
                            Name ="Label352"
                            Caption ="Buy Out $"
                            GroupTable =1
                            LayoutCachedLeft =4404
                            LayoutCachedTop =960
                            LayoutCachedWidth =5628
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
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
                    BorderColor =6108695
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
                    Left =5700
                    Top =600
                    Width =1224
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Book Value"
                    GroupTable =1
                    LayoutCachedLeft =5700
                    LayoutCachedTop =600
                    LayoutCachedWidth =6924
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
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =12348
                    Top =600
                    Width =1236
                    Height =285
                    Name ="Label187"
                    Caption ="FMV Base"
                    GroupTable =1
                    LayoutCachedLeft =12348
                    LayoutCachedTop =600
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =18600
                    Top =600
                    Width =900
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Broker %"
                    GroupTable =1
                    LayoutCachedLeft =18600
                    LayoutCachedTop =600
                    LayoutCachedWidth =19500
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
                    Left =17532
                    Top =600
                    Width =996
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Discount"
                    GroupTable =1
                    LayoutCachedLeft =17532
                    LayoutCachedTop =600
                    LayoutCachedWidth =18528
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17532
                    Top =300
                    Width =996
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="DiscountSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17532
                    LayoutCachedTop =300
                    LayoutCachedWidth =18528
                    LayoutCachedHeight =540
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10040879
                    BorderColor =6108695
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedColor =10040879
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17532
                    Top =960
                    Width =996
                    Height =312
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =15
                    Name ="DiscountCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =17532
                    LayoutCachedTop =960
                    LayoutCachedWidth =18528
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =16272
                            Top =960
                            Width =1188
                            Height =312
                            ForeColor =16777215
                            Name ="Label615"
                            Caption ="Net $"
                            GroupTable =1
                            LayoutCachedLeft =16272
                            LayoutCachedTop =960
                            LayoutCachedWidth =17460
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =13
                            ColumnEnd =13
                            LayoutGroup =1
                            BackThemeColorIndex =1
                            BackShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14964
                    Top =960
                    Width =1236
                    Height =312
                    FontSize =10
                    TabIndex =14
                    BackColor =16317438
                    Name ="TargetSaleCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =14964
                    LayoutCachedTop =960
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =14964
                    Top =600
                    Width =1236
                    Height =285
                    Name ="Label271"
                    Caption ="Target Sale"
                    GroupTable =1
                    LayoutCachedLeft =14964
                    LayoutCachedTop =600
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14964
                    Top =300
                    Width =1236
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="TargetSaleSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14964
                    LayoutCachedTop =300
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =540
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =6108695
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13656
                    Top =300
                    Width =1236
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="FMVValueSubmit"
                    Caption ="▼Submit"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13656
                    LayoutCachedTop =300
                    LayoutCachedWidth =14892
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
                    BorderColor =6108695
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13656
                    Top =960
                    Width =1236
                    Height =312
                    FontSize =10
                    TabIndex =13
                    BackColor =16317438
                    Name ="FMVValueCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =13656
                    LayoutCachedTop =960
                    LayoutCachedWidth =14892
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =13656
                    Top =600
                    Width =1236
                    Height =285
                    Name ="Label498"
                    Caption ="FMV Value"
                    GroupTable =1
                    LayoutCachedLeft =13656
                    LayoutCachedTop =600
                    LayoutCachedWidth =14892
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4404
                    Top =300
                    Width =1224
                    Name ="EmptyCell512"
                    GroupTable =1
                    LayoutCachedLeft =4404
                    LayoutCachedTop =300
                    LayoutCachedWidth =5628
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4404
                    Top =600
                    Width =1224
                    Height =285
                    Name ="EmptyCell513"
                    GroupTable =1
                    LayoutCachedLeft =4404
                    LayoutCachedTop =600
                    LayoutCachedWidth =5628
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9660
                    Top =300
                    Width =2616
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="TermPenaltyCalculate"
                    Caption ="▼Calculate"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Calculate Penalty based on Sold/ Buy/ Book Values, will overwrite exisiting valu"
                        "es."
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9660
                    LayoutCachedTop =300
                    LayoutCachedWidth =12276
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10040879
                    BorderColor =6108695
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedColor =10040879
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =9660
                    Top =600
                    Width =2616
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label559"
                    Caption ="┌─ Term Penalty ─┐"
                    GroupTable =1
                    LayoutCachedLeft =9660
                    LayoutCachedTop =600
                    LayoutCachedWidth =12276
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =9660
                    Top =960
                    Width =1272
                    Height =312
                    BackColor =1643706
                    Name ="Label575"
                    Caption ="Projected"
                    GroupTable =1
                    LayoutCachedLeft =9660
                    LayoutCachedTop =960
                    LayoutCachedWidth =10932
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =87
                    Left =4823
                    Top =187
                    Width =0
                    Height =619
                    BorderColor =10040879
                    Name ="Line582"
                    LayoutCachedLeft =4823
                    LayoutCachedTop =187
                    LayoutCachedWidth =4823
                    LayoutCachedHeight =806
                    BorderThemeColorIndex =-1
                End
                Begin EmptyCell
                    Left =6984
                    Top =300
                    Width =1224
                    Name ="EmptyCell586"
                    GroupTable =1
                    LayoutCachedLeft =6984
                    LayoutCachedTop =300
                    LayoutCachedWidth =8208
                    LayoutCachedHeight =540
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6984
                    Top =600
                    Width =1224
                    Height =285
                    Name ="EmptyCell587"
                    GroupTable =1
                    LayoutCachedLeft =6984
                    LayoutCachedTop =600
                    LayoutCachedWidth =8208
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =12600
                    Top =60
                    Width =3180
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label592"
                    Caption ="┌────────── Unit Valuation ─────────┐                "
                    LayoutCachedLeft =12600
                    LayoutCachedTop =60
                    LayoutCachedWidth =15780
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8280
                    Top =300
                    Width =1308
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ProjectedLPAFLossSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8280
                    LayoutCachedTop =300
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =540
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =13209
                    OldBorderStyle =0
                    BorderColor =12566463
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8280
                    Top =960
                    Width =1308
                    Height =312
                    FontSize =10
                    TabIndex =11
                    BackColor =16317438
                    Name ="ProjectedLPAFLossCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =8280
                    LayoutCachedTop =960
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =6984
                            Top =960
                            Width =1224
                            Height =312
                            BackColor =1643706
                            Name ="Label590"
                            Caption ="Sold Price"
                            GroupTable =1
                            LayoutCachedLeft =6984
                            LayoutCachedTop =960
                            LayoutCachedWidth =8208
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =8280
                    Top =600
                    Width =1308
                    Height =285
                    BackColor =13209
                    BorderColor =12566463
                    Name ="Label595"
                    Caption ="Term Penalty"
                    GroupTable =1
                    LayoutCachedLeft =8280
                    LayoutCachedTop =600
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =16272
                    Top =600
                    Width =1188
                    Height =285
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label605"
                    Caption ="Bank Share"
                    GroupTable =1
                    LayoutCachedLeft =16272
                    LayoutCachedTop =600
                    LayoutCachedWidth =17460
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
                Begin EmptyCell
                    Left =16272
                    Top =300
                    Width =1188
                    Name ="EmptyCell607"
                    GroupTable =1
                    LayoutCachedLeft =16272
                    LayoutCachedTop =300
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =540
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
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
                    Left =12348
                    Top =36
                    Width =1236
                    Height =252
                    FontSize =10
                    TabIndex =11
                    BackColor =16317438
                    Name ="WhSValue"
                    ControlSource ="WhSValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationText ="BO Date is required"
                    ControlTipText ="'Wholesale' value of this Unit"
                    GroupTable =1

                    LayoutCachedLeft =12348
                    LayoutCachedTop =36
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5700
                    Top =36
                    Width =1224
                    Height =252
                    FontSize =10
                    TabIndex =5
                    Name ="AdjBookValue"
                    ControlSource ="AdjBookValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Book Value of the Unit - could differ from Buy Out Value."
                    GroupTable =1

                    LayoutCachedLeft =5700
                    LayoutCachedTop =36
                    LayoutCachedWidth =6924
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
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17532
                    Top =36
                    Width =996
                    Height =252
                    FontSize =10
                    TabIndex =14
                    BackColor =13233663
                    Name ="Discount"
                    ControlSource ="Discount"
                    Format ="$#,##0.00;($#,##0.00)"
                    DefaultValue ="0"
                    ControlTipText ="Value that this unit was discounted"
                    GroupTable =1

                    LayoutCachedLeft =17532
                    LayoutCachedTop =36
                    LayoutCachedWidth =18528
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
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
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18600
                    Top =36
                    Width =900
                    Height =252
                    FontSize =10
                    TabIndex =15
                    Name ="FABrkComm"
                    ControlSource ="FABrkComm"
                    Format ="Percent"
                    ControlTipText ="The percentage to be paid to a outside broker"
                    GroupTable =1

                    LayoutCachedLeft =18600
                    LayoutCachedTop =36
                    LayoutCachedWidth =19500
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =315
                    Width =19425
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
                    LayoutCachedWidth =19425
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
                    Left =14964
                    Top =36
                    Width =1236
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BackColor =16317438
                    Name ="TargetSale"
                    ControlSource ="TargetSale"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Target Retail Sale value of this Unit"
                    GroupTable =1

                    LayoutCachedLeft =14964
                    LayoutCachedTop =36
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13656
                    Top =36
                    Width =1236
                    Height =252
                    FontSize =10
                    TabIndex =10
                    BackColor =16317438
                    Name ="FMVValue"
                    ControlSource ="FMVValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    DefaultValue ="0"
                    ControlTipText ="Fair Market Value of this Unit"
                    GroupTable =1

                    LayoutCachedLeft =13656
                    LayoutCachedTop =36
                    LayoutCachedWidth =14892
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =36
                    Width =1224
                    Height =252
                    FontSize =10
                    TabIndex =4
                    Name ="BuyOut"
                    ControlSource ="BuyOutCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Buy Out Value paid to Assignee Bank"
                    GroupTable =1

                    LayoutCachedLeft =4404
                    LayoutCachedTop =36
                    LayoutCachedWidth =5628
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11004
                    Top =36
                    Width =1272
                    Height =252
                    FontSize =10
                    TabIndex =9
                    Name ="GainLoss"
                    ControlSource ="GainLoss"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Term Penalty, A 'projected' value until the units sales, then it's Buy-Out/Book "
                        "Value vs Actual Sold Price for this unit."
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1

                    LayoutCachedLeft =11004
                    LayoutCachedTop =36
                    LayoutCachedWidth =12276
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9660
                    Top =36
                    Width =1272
                    Height =252
                    FontSize =10
                    TabIndex =8
                    BackColor =15527148
                    Name ="ProjPenalty"
                    ControlSource ="=IIf([TargetSale]>0 And [AdjBookValue]>0,[TargetSale]-[AdjBookValue],IIf([Target"
                        "Sale]>0 And [AdjBookValue]<=0,[TargetSale]-[BuyOutCost],0))"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Projected Penalty, Based on Targeted Sale vs Book Value for this unit."
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1

                    LayoutCachedLeft =9660
                    LayoutCachedTop =36
                    LayoutCachedWidth =10932
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6984
                    Top =36
                    Width =1224
                    Height =252
                    FontSize =10
                    TabIndex =6
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="This is the ACTUAL sold price for this unit (can be edited elsewhere on SOLD val"
                        "ues tab)."
                    GroupTable =1

                    LayoutCachedLeft =6984
                    LayoutCachedTop =36
                    LayoutCachedWidth =8208
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8280
                    Top =36
                    Width =1308
                    Height =252
                    FontSize =10
                    TabIndex =7
                    BackColor =16317438
                    Name ="TermPenalty"
                    ControlSource ="TermPenalty"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationText ="BO Date is required"
                    ControlTipText ="Projected Loss (@LPAF) to remarket this unit. ONCE ENTERED, THIS VALUE WILL NOT "
                        "CHANGE."
                    GroupTable =1

                    LayoutCachedLeft =8280
                    LayoutCachedTop =36
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16272
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =13
                    BackColor =16317438
                    Name ="Text613"
                    ControlSource ="BankShareNetProceeds"
                    Format ="$#,##0.00;($#,##0.00)"
                    DefaultValue ="0"
                    ControlTipText ="Value that this unit was discounted"
                    GroupTable =1

                    LayoutCachedLeft =16272
                    LayoutCachedTop =36
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
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
' See "ReMarketingMassMod_sub_FMVPenalty.cls"
