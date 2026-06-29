Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =95
    Left =12345
    Top =3555
    Right =28500
    Bottom =17460
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
    End
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =735
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4560
                    Top =300
                    Width =6600
                    Height =360
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    Name ="BuyNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =300
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4560
                            Top =120
                            Width =1800
                            Height =180
                            FontSize =8
                            Name ="Label13"
                            Caption ="You're Working with:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =4560
                            LayoutCachedTop =120
                            LayoutCachedWidth =6360
                            LayoutCachedHeight =300
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7890
                    Top =60
                    Width =480
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =7890
                    LayoutCachedTop =60
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =270
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =7395
                    Top =75
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label49"
                    Caption ="Units"
                    LayoutCachedLeft =7395
                    LayoutCachedTop =75
                    LayoutCachedWidth =7845
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7875
                    Top =285
                    Width =480
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="Text63"
                    ControlSource ="=[Units_List].[ItemsSelected].[Count]"

                    LayoutCachedLeft =7875
                    LayoutCachedTop =285
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =495
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =7380
                    Top =300
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label64"
                    Caption ="Units"
                    LayoutCachedLeft =7380
                    LayoutCachedTop =300
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =495
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =540
                    Width =12960
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =735
                    TabIndex =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =255
                    Top =60
                    Width =4560
                    Height =600
                    FontSize =20
                    Name ="Label31"
                    Caption ="BuyOut Transaction Form"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =247
                    Left =30
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8872517
                    Name ="Label294"
                    Caption ="Buyers Order"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =30
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =240
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =12075
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =3
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12075
                    LayoutCachedTop =180
                    LayoutCachedWidth =12825
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =11520
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =4
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

                    LayoutCachedLeft =11520
                    LayoutCachedTop =180
                    LayoutCachedWidth =12075
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8580
                    Top =180
                    Width =480
                    Height =210
                    FontSize =8
                    TabIndex =5
                    ForeColor =-2147483615
                    Name ="ClientGroupID"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =180
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =390
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9420
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =240
                    Top =315
                    Width =5040
                    Height =270
                    FontSize =8
                    Name ="Label89"
                    Caption ="┌────┬──Search ───────────┬──────────┐"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =240
                    LayoutCachedTop =315
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =585
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Top =525
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =8
                    BorderColor =967423
                    Name ="SearchBO"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =525
                    LayoutCachedWidth =720
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5970
                    Top =1740
                    Width =5472
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label58"
                    Caption ="┌───┬  Search ────────┬──────┬────┬────────┐"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =5970
                    LayoutCachedTop =1740
                    LayoutCachedWidth =11442
                    LayoutCachedHeight =2025
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Left =9795
                    Top =1770
                    Width =691
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearLabel"
                    Caption ="◄Clear►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =9795
                    LayoutCachedTop =1770
                    LayoutCachedWidth =10486
                    LayoutCachedHeight =1995
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =5880
                    Top =2250
                    Width =7020
                    Height =6930
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BoundColumn =1
                    BackColor =15527148
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MY"
                        "ear From vw_BuyersOrder Where InvoiceRequestID = 2391 Order BY InvoiceRequestID "
                        "DESC; "
                    ColumnWidths ="0;720;720;2016;288;1152;720;720;0;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5880
                    LayoutCachedTop =2250
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =9180
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6612
                    Top =1980
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchUnit"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6612
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7332
                    LayoutCachedHeight =2235
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7344
                    Top =1980
                    Width =2016
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7344
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2235
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5880
                    Top =1980
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchAsset"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5880
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =2235
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9372
                    Top =1980
                    Width =288
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9372
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =2235
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10836
                    Top =1980
                    Width =828
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchType"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10836
                    LayoutCachedTop =1980
                    LayoutCachedWidth =11664
                    LayoutCachedHeight =2235
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =60
                    Top =840
                    Width =5760
                    Height =8325
                    FontSize =9
                    Name ="BuyersOrderList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InvoiceRequestID as BO ,BuyerShortName AS Buyer ,InvoicePaidDate AS Paid "
                        ",SP ,Count(UnitID) as Units FROM vw_BuyersOrder WHERE WON = True  GROUP BY Invoi"
                        "ceRequestID, BuyerShortName,InvoicePaidDate,SP ORDER BY InvoiceRequestID DESC;"
                    ColumnWidths ="720;2592;1152;576;1152"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =840
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =9165
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =5100
                    Top =540
                    Width =600
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearClientSearch"
                    Caption ="◄Clear"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =540
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =765
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9672
                    Top =1980
                    Width =1152
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchMake"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9672
                    LayoutCachedTop =1980
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =2235
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Top =525
                    Width =2592
                    Height =270
                    FontSize =9
                    TabIndex =9
                    BorderColor =967423
                    Name ="SearchBuyer"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =720
                    LayoutCachedTop =525
                    LayoutCachedWidth =3312
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3312
                    Top =525
                    Width =1152
                    Height =270
                    FontSize =9
                    TabIndex =10
                    BorderColor =967423
                    Name ="SearchPaid"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =3312
                    LayoutCachedTop =525
                    LayoutCachedWidth =4464
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4464
                    Top =525
                    Width =576
                    Height =270
                    FontSize =9
                    TabIndex =11
                    BorderColor =967423
                    Name ="SearchSP"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =4464
                    LayoutCachedTop =525
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Width =4575
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Select A Buyers Order "
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Top =240
                    Width =6187
                    Height =420
                    FontSize =14
                    FontWeight =700
                    TabIndex =12
                    Name ="MakeBOBTF"
                    Caption ="BTF For All Units in Buyers Order"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6480
                    LayoutCachedTop =240
                    LayoutCachedWidth =12667
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12195
                    Top =9120
                    Width =705
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="UnitID"

                    LayoutCachedLeft =12195
                    LayoutCachedTop =9120
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =9375
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =11550
                            Top =9120
                            Width =645
                            Height =255
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label19"
                            Caption ="UnitID:"
                            LayoutCachedLeft =11550
                            LayoutCachedTop =9120
                            LayoutCachedWidth =12195
                            LayoutCachedHeight =9375
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Top =840
                    Width =6187
                    Height =420
                    FontSize =14
                    FontWeight =700
                    TabIndex =14
                    Name ="MakeUnitBTF"
                    Caption ="BTF For Selected Unit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6480
                    LayoutCachedTop =840
                    LayoutCachedWidth =12667
                    LayoutCachedHeight =1260
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BTF_BuyersOrder.cls"
