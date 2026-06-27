Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19140
    DatasheetFontHeight =11
    ItemSuffix =1035
    Left =2805
    Top =4275
    Right =22725
    Bottom =17205
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xea3af13335cee540
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Select * From Units Where GroupRef=0"
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
            Height =1605
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =10800
                    Top =300
                    Width =2475
                    Height =270
                    ForeColor =1643706
                    Name ="Label983"
                    Caption ="┌──────── Per Diem ────────┐"
                    LayoutCachedLeft =10800
                    LayoutCachedTop =300
                    LayoutCachedWidth =13275
                    LayoutCachedHeight =570
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =14790
                    Top =885
                    Width =1665
                    Height =300
                    ForeColor =1643706
                    Name ="Label984"
                    Caption ="┌───  Rent  ───┐"
                    LayoutCachedLeft =14790
                    LayoutCachedTop =885
                    LayoutCachedWidth =16455
                    LayoutCachedHeight =1185
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Left =60
                    Top =1245
                    Width =19080
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =1245
                    LayoutCachedWidth =19140
                    LayoutCachedHeight =1440
                    TabIndex =15
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Width =5430
                    Height =540
                    FontSize =20
                    Name ="Label114"
                    Caption ="Mini Mass Mod- PerDiem Rents"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7635
                    Top =810
                    Width =1065
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="Use_VendorPaid"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7635
                    LayoutCachedTop =810
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1035
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
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8985
                    Top =1095
                    Width =1095
                    Height =240
                    Name ="Label127"
                    Caption ="In Service"
                    RightPadding =38
                    LayoutCachedLeft =8985
                    LayoutCachedTop =1095
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11805
                    Top =1095
                    Width =780
                    Height =240
                    FontWeight =400
                    Name ="Label129"
                    Caption ="PD Rent"
                    RightPadding =38
                    LayoutCachedLeft =11805
                    LayoutCachedTop =1095
                    LayoutCachedWidth =12585
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =14670
                    Top =1095
                    Width =900
                    Height =240
                    FontWeight =400
                    Name ="Label130"
                    Caption ="Schedule"
                    RightPadding =38
                    LayoutCachedLeft =14670
                    LayoutCachedTop =1095
                    LayoutCachedWidth =15570
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =795
                    Top =1095
                    Width =885
                    Height =240
                    FontWeight =400
                    Name ="Label134"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =795
                    LayoutCachedTop =1095
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =7635
                    Top =1095
                    Width =1065
                    Height =270
                    FontWeight =400
                    Name ="Label135"
                    Caption ="Vendor Paid"
                    RightPadding =38
                    LayoutCachedLeft =7635
                    LayoutCachedTop =1095
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1365
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
                    Top =1095
                    Width =690
                    Height =240
                    FontWeight =400
                    Name ="Label132"
                    Caption ="♦Asset:"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =1095
                    LayoutCachedWidth =720
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =1755
                    Top =1095
                    Width =1785
                    Height =240
                    FontWeight =400
                    Name ="Label133"
                    Caption ="VIN"
                    RightPadding =38
                    LayoutCachedLeft =1755
                    LayoutCachedTop =1095
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8970
                    Top =810
                    Width =1095
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="Use_InService"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8970
                    LayoutCachedTop =810
                    LayoutCachedWidth =10065
                    LayoutCachedHeight =1035
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =11835
                    Top =480
                    Width =720
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="PDRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11835
                    LayoutCachedTop =480
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =705
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
                    Left =11835
                    Top =780
                    Width =720
                    Height =255
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =6
                    Name ="PDRentCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =11835
                    LayoutCachedTop =780
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =1035
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6300
                    Top =1095
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label172"
                    Caption ="Delivery"
                    RightPadding =38
                    LayoutCachedLeft =6300
                    LayoutCachedTop =1095
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6285
                    Top =810
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="Use_ActualDelDate"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6285
                    LayoutCachedTop =810
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1035
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
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5025
                    Top =1095
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label184"
                    Caption ="Accepted"
                    RightPadding =38
                    LayoutCachedLeft =5025
                    LayoutCachedTop =1095
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5025
                    Top =810
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="Use_AcceptedDate"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5025
                    LayoutCachedTop =810
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1035
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
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =15630
                    Top =1095
                    Width =900
                    Height =240
                    BackColor =12566463
                    Name ="Label825"
                    Caption ="SYND"
                    RightPadding =38
                    LayoutCachedLeft =15630
                    LayoutCachedTop =1095
                    LayoutCachedWidth =16530
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10710
                    Top =1095
                    Width =1035
                    Height =240
                    FontWeight =400
                    Name ="Label857"
                    Caption ="Start Date"
                    RightPadding =38
                    LayoutCachedLeft =10710
                    LayoutCachedTop =1095
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10710
                    Top =480
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitPDStartDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10710
                    LayoutCachedTop =480
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =705
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
                    Left =10710
                    Top =780
                    Width =1035
                    Height =255
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =5
                    Name ="PDStartDateCBO"
                    Format ="Short Date"
                    StatusBarText ="UNIT TYPE"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =10710
                    LayoutCachedTop =780
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =1035
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =6780
                    Width =645
                    Height =105
                    ColumnOrder =2
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =105
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =3585
                    Top =1080
                    Width =570
                    Height =255
                    FontWeight =400
                    Name ="Label189"
                    Caption ="Status"
                    LayoutCachedLeft =3585
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =1335
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12660
                    Top =1095
                    Width =765
                    Height =270
                    FontWeight =400
                    Name ="Label1001"
                    Caption ="Surrendr"
                    RightPadding =38
                    LayoutCachedLeft =12660
                    LayoutCachedTop =1095
                    LayoutCachedWidth =13425
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12660
                    Top =780
                    Width =720
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =7
                    BackColor =13952764
                    Name ="PDSurrenderCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =12660
                    LayoutCachedTop =780
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1035
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12660
                    Top =480
                    Width =720
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="PDSurrenderSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =480
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =705
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6780
                    Top =225
                    Width =645
                    Height =105
                    FontSize =8
                    TabIndex =1
                    Name ="GroupId"
                    ControlSource ="GroupREF"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =225
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =4920
                    Height =720
                    FontSize =10
                    TabIndex =12
                    ForeColor =12349952
                    Name ="SchNo"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =780
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =18090
                    Top =300
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =13
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =18090
                    LayoutCachedTop =300
                    LayoutCachedWidth =18840
                    LayoutCachedHeight =528
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
                    Left =17520
                    Top =300
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =14
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

                    LayoutCachedLeft =17520
                    LayoutCachedTop =300
                    LayoutCachedWidth =18075
                    LayoutCachedHeight =528
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
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =16845
                    Top =1095
                    Width =1035
                    Height =270
                    FontWeight =400
                    Name ="Label1030"
                    Caption ="AppOut"
                    RightPadding =38
                    LayoutCachedLeft =16845
                    LayoutCachedTop =1095
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =17940
                    Top =1095
                    Width =1035
                    Height =270
                    FontWeight =400
                    Name ="Label1032"
                    Caption ="App IN"
                    RightPadding =38
                    LayoutCachedLeft =17940
                    LayoutCachedTop =1095
                    LayoutCachedWidth =18975
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17220
                    Top =885
                    Width =1305
                    Height =270
                    ForeColor =1643706
                    Name ="Label1034"
                    Caption ="┌───  Title  ───┐"
                    LayoutCachedLeft =17220
                    LayoutCachedTop =885
                    LayoutCachedWidth =18525
                    LayoutCachedHeight =1155
                End
            End
        End
        Begin Section
            Height =387
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10680
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =7
                    Name ="UnitPDStartDate"
                    ControlSource ="UnitPDStartDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =10680
                    LayoutCachedTop =45
                    LayoutCachedWidth =11715
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15600
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =11
                    Name ="UNITRENT"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =15600
                    LayoutCachedTop =45
                    LayoutCachedWidth =16500
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14640
                    Top =45
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =10
                    Name ="UNITSRENT"
                    ControlSource ="UNITSRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =14640
                    LayoutCachedTop =45
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7605
                    Top =45
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =13
                    Name ="UnitVendorPytDate"
                    ControlSource ="UnitVendorPytDate"
                    Format ="Short Date"
                    StatusBarText ="unit make"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =7605
                    LayoutCachedTop =45
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6270
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="UnitActualDelDate"
                    ControlSource ="UnitActualDelDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =6270
                    LayoutCachedTop =45
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4995
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="UnitAcceptDate"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =45
                    LayoutCachedWidth =6030
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
                    Left =8955
                    Top =45
                    Width =1095
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="UnitInServ"
                    ControlSource ="UnitInServ"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =8955
                    LayoutCachedTop =45
                    LayoutCachedWidth =10050
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
                    Left =11805
                    Top =45
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="UNITPDRENT"
                    ControlSource ="UNITPDRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =11805
                    LayoutCachedTop =45
                    LayoutCachedWidth =12525
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
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
                    Locked = NotDefault
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
                    TabIndex =3
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
                    OverlapFlags =95
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =19080
                    Height =72
                    TabIndex =12
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
                    LayoutCachedWidth =19080
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
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12630
                    Top =45
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =9
                    BackColor =13952764
                    Name ="PerDiemSurrender"
                    ControlSource ="PerDiemSurrender"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =12630
                    LayoutCachedTop =45
                    LayoutCachedWidth =13350
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13860
                    Top =60
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =14
                    Name ="Text1029"
                    ControlSource ="=Round([UNITSRENT]/30,2)"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"

                    LayoutCachedLeft =13860
                    LayoutCachedTop =60
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16845
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =15
                    Name ="Text1031"
                    ControlSource ="tSent"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =16845
                    LayoutCachedTop =45
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17940
                    Top =45
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =16
                    Name ="Text1033"
                    ControlSource ="tBack"
                    StatusBarText ="MONTHLY RENT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =17940
                    LayoutCachedTop =45
                    LayoutCachedWidth =18975
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
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
' See "MassMod_PerdiemRents.cls"
