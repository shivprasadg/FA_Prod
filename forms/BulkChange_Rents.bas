Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21660
    DatasheetFontHeight =11
    ItemSuffix =1007
    Left =4890
    Top =5115
    Right =24435
    Bottom =16185
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x4016b62e1528e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            FontSize =9
            FontWeight =700
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1185
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =18795
                    Width =1695
                    Height =300
                    Name ="Label996"
                    Caption ="┌─── Residual ───┐"
                    LayoutCachedLeft =18795
                    LayoutCachedWidth =20490
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =13350
                    Top =30
                    Width =2115
                    Height =270
                    ForeColor =1643706
                    Name ="Label983"
                    Caption ="┌────── Per Diem ──────┐"
                    LayoutCachedLeft =13350
                    LayoutCachedTop =30
                    LayoutCachedWidth =15465
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =11220
                    Top =30
                    Width =1665
                    Height =300
                    ForeColor =1643706
                    Name ="Label984"
                    Caption ="┌───  Rent  ───┐"
                    LayoutCachedLeft =11220
                    LayoutCachedTop =30
                    LayoutCachedWidth =12885
                    LayoutCachedHeight =330
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Left =30
                    Top =990
                    Width =21600
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =30
                    LayoutCachedTop =990
                    LayoutCachedWidth =21630
                    LayoutCachedHeight =1185
                    TabIndex =32
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Width =4110
                    Height =510
                    FontSize =20
                    Name ="Label114"
                    Caption ="Dates|Cost |Rent | Resi"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =510
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5325
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="EstDeliveryDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38

                    LayoutCachedLeft =5325
                    LayoutCachedTop =225
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =450
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4215
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="BuildDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4215
                    LayoutCachedTop =225
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =450
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8625
                    Top =840
                    Width =1095
                    Height =240
                    Name ="Label127"
                    Caption ="In Service"
                    RightPadding =38
                    LayoutCachedLeft =8625
                    LayoutCachedTop =840
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5325
                    Top =840
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label128"
                    Caption ="Est. Delivery"
                    RightPadding =38
                    LayoutCachedLeft =5325
                    LayoutCachedTop =840
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14130
                    Top =840
                    Width =780
                    Height =240
                    FontWeight =400
                    Name ="Label129"
                    Caption ="PD Rent"
                    RightPadding =38
                    LayoutCachedLeft =14130
                    LayoutCachedTop =840
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11115
                    Top =840
                    Width =900
                    Height =240
                    FontWeight =400
                    Name ="Label130"
                    Caption ="Schedule"
                    RightPadding =38
                    LayoutCachedLeft =11115
                    LayoutCachedTop =840
                    LayoutCachedWidth =12015
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    Left =15720
                    Top =840
                    Width =900
                    Height =240
                    FontWeight =400
                    Name ="Label131"
                    Caption ="1st ExtRent:"
                    RightPadding =38
                    LayoutCachedLeft =15720
                    LayoutCachedTop =840
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =18
                    ColumnEnd =18
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =795
                    Top =840
                    Width =885
                    Height =240
                    FontWeight =400
                    Name ="Label134"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =795
                    LayoutCachedTop =840
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4215
                    Top =840
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label135"
                    Caption ="Build Date"
                    RightPadding =38
                    LayoutCachedLeft =4215
                    LayoutCachedTop =840
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =30
                    Top =840
                    Width =690
                    Height =240
                    FontWeight =400
                    Name ="Label132"
                    Caption ="♦Asset:"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =840
                    LayoutCachedWidth =720
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =1755
                    Top =840
                    Width =1785
                    Height =240
                    FontWeight =400
                    Name ="Label133"
                    Caption ="VIN"
                    RightPadding =38
                    LayoutCachedLeft =1755
                    LayoutCachedTop =840
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8655
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="InServiceSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8655
                    LayoutCachedTop =225
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =450
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8655
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =17
                    Name ="InservDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =8655
                    LayoutCachedTop =525
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14160
                    Top =225
                    Width =720
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="PDRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14160
                    LayoutCachedTop =225
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =450
                    ColumnStart =17
                    ColumnEnd =17
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =9
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =9
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14160
                    Top =525
                    Width =720
                    Height =255
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =23
                    Name ="PDRentCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =14160
                    LayoutCachedTop =525
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11115
                    Top =525
                    Width =900
                    Height =255
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =19
                    Name ="StripCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =11115
                    LayoutCachedTop =525
                    LayoutCachedWidth =12015
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15720
                    Top =225
                    Width =900
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="FirstExtRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15720
                    LayoutCachedTop =225
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =450
                    ColumnStart =18
                    ColumnEnd =18
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15720
                    Top =525
                    Width =900
                    Height =255
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =22
                    Name ="FirstExtCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =15720
                    LayoutCachedTop =525
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9885
                    Top =840
                    Width =1155
                    Height =240
                    FontWeight =400
                    BackColor =14151142
                    Name ="Label163"
                    Caption ="Unit Cost"
                    RightPadding =38
                    LayoutCachedLeft =9885
                    LayoutCachedTop =840
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9885
                    Top =225
                    Width =1155
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="CostSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9885
                    LayoutCachedTop =225
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =450
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9885
                    Top =525
                    Width =1155
                    Height =255
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =18
                    Name ="CostCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =9885
                    LayoutCachedTop =525
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4215
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =13
                    Name ="BuildDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =4215
                    LayoutCachedTop =525
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5325
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =14
                    Name ="EstDeliveryDate"
                    Format ="Short Date"
                    StatusBarText ="UNIT TYPE"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =5325
                    LayoutCachedTop =525
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6435
                    Top =840
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label172"
                    Caption ="Delivery"
                    RightPadding =38
                    LayoutCachedLeft =6435
                    LayoutCachedTop =840
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6435
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ActualDelDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6435
                    LayoutCachedTop =225
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =450
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6435
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =15
                    Name ="ActualDelDate"
                    Format ="Short Date"
                    StatusBarText ="UNIT TYPE"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =6435
                    LayoutCachedTop =525
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =17625
                    Top =840
                    Width =825
                    Height =240
                    FontWeight =400
                    BackColor =14151142
                    Name ="Label177"
                    Caption ="Tax"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =17625
                    LayoutCachedTop =840
                    LayoutCachedWidth =18450
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =20
                    ColumnEnd =20
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17625
                    Top =225
                    Width =825
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="UnitTaxSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="This will calulate the Unit's Taxes @ the Local Rate and if Exempt."
                    LeftPadding =0
                    RightPadding =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17625
                    LayoutCachedTop =225
                    LayoutCachedWidth =18450
                    LayoutCachedHeight =450
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7545
                    Top =840
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label184"
                    Caption ="Accepted"
                    RightPadding =38
                    LayoutCachedLeft =7545
                    LayoutCachedTop =840
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7545
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="AcceptedDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7545
                    LayoutCachedTop =225
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =450
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7545
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =16
                    Name ="AcceptedDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =7545
                    LayoutCachedTop =525
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16470
                    Top =840
                    Width =0
                    Height =240
                    FontWeight =400
                    BackColor =14151142
                    Name ="Label201"
                    Caption ="UnitDocAccpt"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =16470
                    LayoutCachedTop =840
                    LayoutCachedWidth =16470
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =19
                    ColumnEnd =19
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12075
                    Top =525
                    Width =900
                    Height =255
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =20
                    Name ="SyndRentCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =12075
                    LayoutCachedTop =525
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12075
                    Top =840
                    Width =900
                    Height =240
                    BackColor =12566463
                    Name ="Label825"
                    Caption ="SYND"
                    RightPadding =38
                    LayoutCachedLeft =12075
                    LayoutCachedTop =840
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =17625
                    Top =525
                    Width =825
                    Height =255
                    FontSize =10
                    FontWeight =400
                    ForeColor =16777215
                    Name ="Label854"
                    Caption ="Taxes"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =17625
                    LayoutCachedTop =525
                    LayoutCachedWidth =18450
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =20
                    ColumnEnd =20
                    BackThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13065
                    Top =840
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label857"
                    Caption ="Start Date"
                    RightPadding =38
                    LayoutCachedLeft =13065
                    LayoutCachedTop =840
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13065
                    Top =225
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="UnitPDStartDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13065
                    LayoutCachedTop =225
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =450
                    ColumnStart =16
                    ColumnEnd =16
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =9
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =9
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13065
                    Top =525
                    Width =1035
                    Height =255
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =21
                    Name ="PDStartDateCBO"
                    Format ="Short Date"
                    StatusBarText ="UNIT TYPE"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =13065
                    LayoutCachedTop =525
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =9270
                    Top =60
                    Width =645
                    Height =105
                    ColumnOrder =2
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =9270
                    LayoutCachedTop =60
                    LayoutCachedWidth =9915
                    LayoutCachedHeight =165
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12075
                    Top =225
                    Width =900
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="SyndRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12075
                    LayoutCachedTop =225
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =450
                    ColumnStart =15
                    ColumnEnd =15
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =1643706
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11115
                    Top =225
                    Width =900
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="StripSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11115
                    LayoutCachedTop =225
                    LayoutCachedWidth =12015
                    LayoutCachedHeight =450
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =2366701
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =2366701
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =3585
                    Top =840
                    Width =570
                    Height =255
                    FontWeight =400
                    Name ="Label189"
                    Caption ="Status"
                    LayoutCachedLeft =3585
                    LayoutCachedTop =840
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =1095
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =18495
                    Top =840
                    Width =1185
                    Height =240
                    FontWeight =400
                    BackColor =14151142
                    Name ="Label988"
                    Caption ="Schedule"
                    RightPadding =38
                    LayoutCachedLeft =18495
                    LayoutCachedTop =840
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =18510
                    Top =225
                    Width =1155
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
                    Name ="SalesResidSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18510
                    LayoutCachedTop =225
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =450
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18510
                    Top =525
                    Width =1155
                    Height =255
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =25
                    Name ="SalesResidCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =18510
                    LayoutCachedTop =525
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19740
                    Top =840
                    Width =1185
                    Height =240
                    BackColor =14151142
                    Name ="Label992"
                    Caption ="SYND"
                    RightPadding =38
                    LayoutCachedLeft =19740
                    LayoutCachedTop =840
                    LayoutCachedWidth =20925
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =19740
                    Top =225
                    Width =1155
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =26
                    Name ="SynResidSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19740
                    LayoutCachedTop =225
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =450
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19740
                    Top =525
                    Width =1155
                    Height =255
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =27
                    BackColor =14151142
                    Name ="SynResidCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =19740
                    LayoutCachedTop =525
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =16665
                    Top =840
                    Width =900
                    Height =240
                    FontWeight =400
                    Name ="Label997"
                    Caption ="MTM Rent:"
                    RightPadding =38
                    LayoutCachedLeft =16665
                    LayoutCachedTop =840
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =18
                    ColumnEnd =18
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16665
                    Top =225
                    Width =900
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =28
                    Name ="MTMRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16665
                    LayoutCachedTop =225
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =450
                    ColumnStart =18
                    ColumnEnd =18
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16665
                    Top =525
                    Width =900
                    Height =255
                    ColumnOrder =17
                    FontSize =9
                    TabIndex =29
                    Name ="MTMRentCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =16665
                    LayoutCachedTop =525
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14955
                    Top =840
                    Width =765
                    Height =270
                    FontWeight =400
                    Name ="Label1001"
                    Caption ="Surrendr"
                    RightPadding =38
                    LayoutCachedLeft =14955
                    LayoutCachedTop =840
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =1110
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14955
                    Top =525
                    Width =720
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =30
                    BackColor =13952764
                    Name ="PDSurrenderCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =14955
                    LayoutCachedTop =525
                    LayoutCachedWidth =15675
                    LayoutCachedHeight =780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14955
                    Top =225
                    Width =720
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =31
                    Name ="PDSurrenderSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14955
                    LayoutCachedTop =225
                    LayoutCachedWidth =15675
                    LayoutCachedHeight =450
                    ColumnStart =17
                    ColumnEnd =17
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =887479
                    BorderThemeColorIndex =1
                    HoverColor =887479
                    PressedColor =887479
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =387
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =21000
                    Top =45
                    Width =420
                    Height =270
                    FontSize =9
                    TabIndex =22
                    Name ="UnitIFAmt"
                    ControlSource ="UnitIFAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =21000
                    LayoutCachedTop =45
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13065
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =11
                    Name ="UnitPDStartDate"
                    ControlSource ="UnitPDStartDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =13065
                    LayoutCachedTop =45
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12075
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =10
                    Name ="UNITRENT"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =12075
                    LayoutCachedTop =45
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11115
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =9
                    Name ="UNITSRENT"
                    ControlSource ="UNITSRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =11115
                    LayoutCachedTop =45
                    LayoutCachedWidth =12015
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17625
                    Top =45
                    Width =825
                    Height =270
                    FontSize =9
                    TabIndex =15
                    Name ="UnitTax"
                    ControlSource ="UnitTax"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17625
                    LayoutCachedTop =45
                    LayoutCachedWidth =18450
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =20
                    ColumnEnd =20
                    BackThemeColorIndex =4
                    BackTint =20.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9885
                    Top =45
                    Width =1155
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="UnitFinalCost"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =9885
                    LayoutCachedTop =45
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4215
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="UnitBuildDate"
                    ControlSource ="UnitBuildDate"
                    Format ="Short Date"
                    StatusBarText ="unit make"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =4215
                    LayoutCachedTop =45
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5325
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="UnitEstDelDate"
                    ControlSource ="UnitEstDelDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =5325
                    LayoutCachedTop =45
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6435
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="UnitActualDelDate"
                    ControlSource ="UnitActualDelDate"
                    Format ="Short Date"
                    OnClick ="[Event Procedure]"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =6435
                    LayoutCachedTop =45
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7545
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="UnitAcceptDate"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =7545
                    LayoutCachedTop =45
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8655
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =7
                    Name ="UnitInServ"
                    ControlSource ="UnitInServ"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =8655
                    LayoutCachedTop =45
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14160
                    Top =45
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =12
                    Name ="UNITPDRENT"
                    ControlSource ="UNITPDRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =14160
                    LayoutCachedTop =45
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15720
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =13
                    Name ="FirstExtRent"
                    ControlSource ="FirstExtRent"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =15720
                    LayoutCachedTop =45
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =18
                    ColumnEnd =18
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =30
                    Top =45
                    Width =690
                    Height =270
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =45
                    LayoutCachedWidth =720
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1755
                    Top =45
                    Width =1785
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =1755
                    LayoutCachedTop =45
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =795
                    Top =45
                    Width =885
                    Height =270
                    FontSize =9
                    TabIndex =1
                    Name ="UnitYr"
                    ControlSource ="UnitUnitNum"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =795
                    LayoutCachedTop =45
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =127
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16470
                    Top =45
                    Width =0
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="UnitDocAccpt"
                    ControlSource ="UnitDocAccpt"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16470
                    LayoutCachedTop =45
                    LayoutCachedWidth =16470
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =19
                    ColumnEnd =19
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3585
                    Top =45
                    Width =570
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"

                    LayoutCachedLeft =3585
                    LayoutCachedTop =45
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =21600
                    Height =72
                    TabIndex =17
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =315
                    LayoutCachedWidth =21600
                    LayoutCachedHeight =387
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18510
                    Top =45
                    Width =1155
                    Height =270
                    FontSize =9
                    TabIndex =18
                    Name ="UnitSalesResid"
                    ControlSource ="UnitSalesResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =18510
                    LayoutCachedTop =45
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19740
                    Top =45
                    Width =1155
                    Height =270
                    FontSize =9
                    TabIndex =19
                    BackColor =14151142
                    Name ="UnitSynResid"
                    ControlSource ="UnitSynResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =19740
                    LayoutCachedTop =45
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16665
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =20
                    Name ="MTMRent"
                    ControlSource ="MTMRent"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =16665
                    LayoutCachedTop =45
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =18
                    ColumnEnd =18
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14955
                    Top =45
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =21
                    BackColor =13952764
                    Name ="PerDiemSurrender"
                    ControlSource ="PerDiemSurrender"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =14955
                    LayoutCachedTop =45
                    LayoutCachedWidth =15675
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
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
' See "BulkChange_Rents.cls"
