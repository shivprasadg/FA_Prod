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
    Width =14460
    DatasheetFontHeight =11
    ItemSuffix =1040
    Left =7320
    Top =6495
    Right =22035
    Bottom =10845
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xd6c15e7790d0e540
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT * FROM Units INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Upda"
        "te_WorkTable.UnitID;"
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
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =6548
                    Top =300
                    Width =2355
                    Height =270
                    ForeColor =1643706
                    Name ="Label983"
                    Caption ="┌──────── Invoice  ────────┐"
                    LayoutCachedLeft =6548
                    LayoutCachedTop =300
                    LayoutCachedWidth =8903
                    LayoutCachedHeight =570
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Left =60
                    Top =1245
                    Width =14400
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
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =1440
                    TabIndex =13
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
                    Caption ="Mini Mass Mod- Invoice/Draw"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =7868
                    Top =1095
                    Width =1095
                    Height =270
                    FontWeight =400
                    Name ="Label129"
                    Caption ="Invoice Date"
                    RightPadding =38
                    LayoutCachedLeft =7868
                    LayoutCachedTop =1095
                    LayoutCachedWidth =8963
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
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
                    OverlapFlags =215
                    Left =7868
                    Top =480
                    Width =1095
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="InvoiceDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7868
                    LayoutCachedTop =480
                    LayoutCachedWidth =8963
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
                    Left =7868
                    Top =780
                    Width =1095
                    Height =255
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =5
                    Name ="UnitInvoiceDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =7868
                    LayoutCachedTop =780
                    LayoutCachedWidth =8963
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
                    TextAlign =2
                    Left =6240
                    Top =1095
                    Width =1515
                    Height =270
                    FontWeight =400
                    Name ="Label857"
                    Caption ="Invoice Number"
                    RightPadding =38
                    LayoutCachedLeft =6240
                    LayoutCachedTop =1095
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =1365
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6240
                    Top =480
                    Width =1515
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitInvoiceNumSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6240
                    LayoutCachedTop =480
                    LayoutCachedWidth =7755
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6240
                    Top =780
                    Width =1515
                    Height =255
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =4
                    Name ="UnitInvoiceNumCBO"
                    StatusBarText ="UNIT TYPE"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =6240
                    LayoutCachedTop =780
                    LayoutCachedWidth =7755
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
                    OverlapFlags =85
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7500
                    Width =645
                    Height =105
                    FontSize =8
                    TabIndex =1
                    Name ="GroupId"
                    ControlSource ="GroupREF"

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =8145
                    LayoutCachedHeight =105
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8280
                    Width =6120
                    Height =360
                    FontSize =10
                    TabIndex =6
                    ForeColor =12349952
                    Name ="SchNo"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =13470
                    Top =420
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =7
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13470
                    LayoutCachedTop =420
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =648
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
                    Left =12900
                    Top =420
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =8
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

                    LayoutCachedLeft =12900
                    LayoutCachedTop =420
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =648
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
                    OverlapFlags =247
                    TextAlign =2
                    Left =9488
                    Top =1095
                    Width =930
                    Height =270
                    FontWeight =400
                    Name ="Label203"
                    Caption ="DRAW"
                    LayoutCachedLeft =9488
                    LayoutCachedTop =1095
                    LayoutCachedWidth =10418
                    LayoutCachedHeight =1365
                    RowStart =2
                    RowEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9488
                    Top =450
                    Width =930
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="DrawDownSubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9488
                    LayoutCachedTop =450
                    LayoutCachedWidth =10418
                    LayoutCachedHeight =705
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
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =9488
                    Top =780
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =10
                    BackColor =15590879
                    Name ="DrawDownCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID GROUP BY DrawDowns.DrawDownId, DrawDowns.DrawNum"
                        "ber, DrawDowns.DrawName, Banks.BankShortName, Banks.BankID, DrawDowns.GroupID HA"
                        "VING (((DrawDowns.GroupID)=[Forms]![MassMod_InvoiceDraw]![GroupId]));"
                    ColumnWidths ="0;432;1728;864;0"
                    Tag ="Medit"

                    LayoutCachedLeft =9488
                    LayoutCachedTop =780
                    LayoutCachedWidth =10418
                    LayoutCachedHeight =1035
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4200
                    Top =1095
                    Width =1785
                    Height =240
                    FontWeight =400
                    Name ="Label1035"
                    Caption ="Vendor Cost"
                    RightPadding =38
                    LayoutCachedLeft =4200
                    LayoutCachedTop =1095
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10980
                    Top =1095
                    Width =1080
                    Height =270
                    FontWeight =400
                    Name ="Label126"
                    Caption ="Vendor Paid"
                    LayoutCachedLeft =10980
                    LayoutCachedTop =1095
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =1365
                    RowStart =2
                    RowEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10980
                    Top =450
                    Width =1080
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="VendorPayDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10980
                    LayoutCachedTop =450
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =705
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10980
                    Top =780
                    Width =1080
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="VendorPayDate"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =10980
                    LayoutCachedTop =780
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =1035
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =342
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10980
                    Top =30
                    Width =1080
                    Height =255
                    FontSize =9
                    TabIndex =9
                    Name ="UnitVendorPytDate"
                    ControlSource ="UnitVendorPytDate"
                    Format ="Short Date"

                    LayoutCachedLeft =10980
                    LayoutCachedTop =30
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4200
                    Top =30
                    Width =1785
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="VendorCost"
                    ControlSource ="VendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="VIN number must be unique in table"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =4200
                    LayoutCachedTop =30
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =9488
                    Top =30
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BackColor =15590879
                    Name ="DrawDownId"
                    ControlSource ="DrawDownId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID GROUP BY DrawDowns.DrawDownId, DrawDowns.DrawNum"
                        "ber, DrawDowns.DrawName, Banks.BankShortName, Banks.BankID, DrawDowns.GroupID HA"
                        "VING (((DrawDowns.GroupID)=[Forms]![MassMod_InvoiceDraw]![GroupId]));"
                    ColumnWidths ="0;432;1728;864;0"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Format ="Standard"

                    LayoutCachedLeft =9488
                    LayoutCachedTop =30
                    LayoutCachedWidth =10418
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3585
                    Top =30
                    Width =570
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"

                    LayoutCachedLeft =3585
                    LayoutCachedTop =30
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =795
                    Top =30
                    Width =885
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="UnitYr"
                    ControlSource ="UnitUnitNum"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =795
                    LayoutCachedTop =30
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1755
                    Top =30
                    Width =1785
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =1755
                    LayoutCachedTop =30
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =690
                    Height =255
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="Units.UnitID"
                    OnDblClick ="[Event Procedure]"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =720
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7868
                    Top =30
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="UnitInvoiceDate"
                    ControlSource ="UnitInvoiceDate"
                    Format ="Short Date"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =7868
                    LayoutCachedTop =30
                    LayoutCachedWidth =8963
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6240
                    Top =30
                    Width =1515
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="UnitInvoiceNum"
                    ControlSource ="UnitInvoiceNum"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =30
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =285
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
                    OverlapFlags =247
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =270
                    Width =14400
                    Height =72
                    TabIndex =6
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =270
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =342
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
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
' See "MassMod_InvoiceDraw.cls"
