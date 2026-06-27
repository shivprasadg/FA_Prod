Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28095
    DatasheetFontHeight =11
    ItemSuffix =135
    Left =2175
    Top =9960
    Right =19725
    Bottom =26850
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x01db86372d6ae440
    End
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
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
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
            Height =765
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11595
                    Width =720
                    Height =270
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =12
                    Name ="ClientGroupID"

                    LayoutCachedLeft =11595
                    LayoutCachedWidth =12315
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9450
                    Width =720
                    Height =270
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =7
                    Name ="MLID"

                    LayoutCachedLeft =9450
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =270
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =570
                    Width =28080
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =570
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =765
                    TabIndex =18
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10170
                    Width =720
                    Height =270
                    ColumnOrder =5
                    FontSize =10
                    Name ="FAID"

                    LayoutCachedLeft =10170
                    LayoutCachedWidth =10890
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10890
                    Width =720
                    Height =270
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =1
                    Name ="ClientID"

                    LayoutCachedLeft =10890
                    LayoutCachedWidth =11610
                    LayoutCachedHeight =270
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =22920
                    Width =1695
                    Height =330
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label32"
                    Caption ="DATA"
                    HorizontalAnchor =2
                    LayoutCachedLeft =22920
                    LayoutCachedWidth =24615
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20250
                    Width =480
                    Height =210
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =5
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =20250
                    LayoutCachedWidth =20730
                    LayoutCachedHeight =210
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =22695
                    Top =150
                    Width =1920
                    Height =480
                    FontSize =24
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label76"
                    Caption ="LIVE EDIT"
                    HorizontalAnchor =2
                    LayoutCachedLeft =22695
                    LayoutCachedTop =150
                    LayoutCachedWidth =24615
                    LayoutCachedHeight =630
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =6870
                    Width =2505
                    Height =288
                    FontSize =9
                    TabIndex =6
                    Name ="UpdatePortal"
                    Caption ="▲ Send Update To Portal! ▲"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6870
                    LayoutCachedWidth =9375
                    LayoutCachedHeight =288
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
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
                    Left =14355
                    Top =240
                    Width =1770
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="CofA_Report"
                    Caption ="Create a C-of-A "
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14355
                    LayoutCachedTop =240
                    LayoutCachedWidth =16125
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
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
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13300
                    Top =270
                    Width =839
                    Height =299
                    ColumnOrder =7
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=[BulkChange_Desc].Form!UnitCount"

                    LayoutCachedLeft =13300
                    LayoutCachedTop =270
                    LayoutCachedWidth =14139
                    LayoutCachedHeight =569
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =13275
                            Top =60
                            Width =900
                            Height =540
                            FontSize =8
                            BackColor =1643706
                            BorderColor =1643706
                            ForeColor =16777215
                            Name ="Label63"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =13275
                            LayoutCachedTop =60
                            LayoutCachedWidth =14175
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20880
                    Width =720
                    Height =180
                    ColumnOrder =8
                    FontSize =8
                    TabIndex =3
                    ForeColor =9211020
                    Name ="GrpID"

                    LayoutCachedLeft =20880
                    LayoutCachedWidth =21600
                    LayoutCachedHeight =180
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3375
                    Top =240
                    Width =5460
                    Height =360
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    Name ="SchNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3375
                    LayoutCachedTop =240
                    LayoutCachedWidth =8835
                    LayoutCachedHeight =600
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3375
                            Top =60
                            Width =1800
                            Height =180
                            FontSize =8
                            Name ="Label13"
                            Caption ="You're Working with:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =3375
                            LayoutCachedTop =60
                            LayoutCachedWidth =5175
                            LayoutCachedHeight =240
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Top =75
                    Width =3375
                    Height =570
                    FontSize =20
                    Name ="Label31"
                    Caption ="Mass Modification"
                    FontName ="Segoe UI"
                    LayoutCachedTop =75
                    LayoutCachedWidth =3375
                    LayoutCachedHeight =645
                End
                Begin Label
                    OverlapFlags =247
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8872517
                    Name ="Label294"
                    Caption ="Operations"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5460
                    Top =255
                    Width =1095
                    Height =255
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =9
                    Name ="PLBAsOfDate"
                    Format ="Short Date"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =255
                    LayoutCachedWidth =6555
                    LayoutCachedHeight =510
                    BorderThemeColorIndex =4
                    BorderShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =5370
                            Width =1365
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label75"
                            Caption ="PLB 'As Of Date'"
                            RightPadding =38
                            LayoutCachedLeft =5370
                            LayoutCachedWidth =6735
                            LayoutCachedHeight =255
                            RowStart =3
                            RowEnd =3
                            ColumnStart =1
                            ColumnEnd =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10935
                    Top =240
                    Width =2100
                    Height =270
                    FontSize =9
                    TabIndex =10
                    Name ="COFLeaseReport"
                    Caption ="Certificate of Lease"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10935
                    LayoutCachedTop =240
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
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
                    OverlapFlags =247
                    Left =8775
                    Top =240
                    Width =2100
                    Height =270
                    FontSize =9
                    TabIndex =11
                    Name ="SUF_Mass"
                    Caption ="SUF for Selected Units"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8775
                    LayoutCachedTop =240
                    LayoutCachedWidth =10875
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
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
                    Left =20490
                    Top =255
                    Width =1710
                    Height =255
                    FontSize =10
                    TabIndex =13
                    Name ="OpenGroupForm"
                    Caption ="Open Group Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20490
                    LayoutCachedTop =255
                    LayoutCachedWidth =22200
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =98559
                    BorderThemeColorIndex =1
                    HoverColor =98559
                    PressedColor =98559
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =26355
                    Top =150
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

                    LayoutCachedLeft =26355
                    LayoutCachedTop =150
                    LayoutCachedWidth =26910
                    LayoutCachedHeight =378
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
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =26926
                    Top =150
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =15
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =26926
                    LayoutCachedTop =150
                    LayoutCachedWidth =27676
                    LayoutCachedHeight =378
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16305
                    Top =225
                    Width =720
                    Height =270
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =16
                    Name ="Text120"

                    LayoutCachedLeft =16305
                    LayoutCachedTop =225
                    LayoutCachedWidth =17025
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7920
                    Top =270
                    Width =405
                    Height =255
                    FontSize =10
                    TabIndex =17
                    Name ="ModMade"
                    Format ="Yes/No"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =270
                    LayoutCachedWidth =8325
                    LayoutCachedHeight =525
                    BorderThemeColorIndex =4
                    BorderShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =7140
                            Top =15
                            Width =1365
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label129"
                            Caption ="PLB 'As Of Date'"
                            RightPadding =38
                            LayoutCachedLeft =7140
                            LayoutCachedTop =15
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =270
                            RowStart =3
                            RowEnd =3
                            ColumnStart =1
                            ColumnEnd =1
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13560
            BackColor =14211288
            Name ="Detail"
            AlternateBackColor =-2147483606
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =13020
                    Width =15000
                    Height =420
                    BorderColor =16777215
                    Name ="Box100"
                    HorizontalAnchor =2
                    LayoutCachedLeft =13020
                    LayoutCachedWidth =28020
                    LayoutCachedHeight =420
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =60
                    Top =180
                    Width =2733
                    FontSize =9
                    BackColor =13434879
                    Name ="SearchClient"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =2793
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =60
                            Width =1950
                            Height =225
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label45"
                            Caption ="Search Client:                       "
                            LayoutCachedLeft =60
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =7125
                    Width =5100
                    Height =6435
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BorderColor =14211288
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="720;1872;720;7200"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =7125
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =13560
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeShade =50.0
                End
                Begin Tab
                    OverlapFlags =223
                    Left =5190
                    Width =22905
                    Height =13515
                    TabIndex =3
                    Name ="TabCtl52"
                    VerticalAnchor =2

                    LayoutCachedLeft =5190
                    LayoutCachedWidth =28095
                    LayoutCachedHeight =13515
                    UseTheme =1
                    BackThemeColorIndex =1
                    OldBorderStyle =0
                    BorderColor =10855845
                    HoverColor =10855845
                    PressedColor =14211288
                    HoverForeThemeColorIndex =6
                    HoverForeShade =50.0
                    PressedForeColor =6108695
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =5265
                            Top =479
                            Width =22761
                            Height =12960
                            BorderColor =0
                            Name ="Description | Spec"
                            OnClick ="[Event Procedure]"
                            EventProcPrefix ="Description___Spec"
                            GridlineColor =0
                            LayoutCachedLeft =5265
                            LayoutCachedTop =479
                            LayoutCachedWidth =28026
                            LayoutCachedHeight =13439
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    Left =5346
                                    Top =479
                                    Width =22680
                                    Height =12960
                                    Name ="BulkChange_Desc"
                                    SourceObject ="Form.BulkChange_Desc"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5346
                                    LayoutCachedTop =479
                                    LayoutCachedWidth =28026
                                    LayoutCachedHeight =13439
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =480
                            Width =22755
                            Height =12960
                            Name ="EngTrans"
                            Caption ="Engine | Trans"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =480
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5346
                                    Top =480
                                    Width =22320
                                    Height =12960
                                    Name ="BulkChange_Desc2"
                                    SourceObject ="Form.BulkChange_Desc2"

                                    LayoutCachedLeft =5346
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =27666
                                    LayoutCachedHeight =13440
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =479
                            Width =22755
                            Height =12961
                            Name ="Rents"
                            Caption ="Dates|Cost |Rent | Residual"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =479
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5346
                                    Top =479
                                    Width =22650
                                    Height =12960
                                    Name ="BulkChange_Rents"
                                    SourceObject ="Form.BulkChange_Rents"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5346
                                    LayoutCachedTop =479
                                    LayoutCachedWidth =27996
                                    LayoutCachedHeight =13439
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =475
                            Width =22755
                            Height =12965
                            BorderColor =0
                            Name ="Vendor"
                            Caption ="Vendor | Invoice | Draw"
                            GridlineColor =0
                            LayoutCachedLeft =5265
                            LayoutCachedTop =475
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5346
                                    Top =475
                                    Width =22560
                                    Height =12960
                                    Name ="BulkChange_Vendor"
                                    SourceObject ="Form.BulkChange_Vendor"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5346
                                    LayoutCachedTop =475
                                    LayoutCachedWidth =27906
                                    LayoutCachedHeight =13435
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =475
                            Width =22755
                            Height =12965
                            BorderColor =0
                            Name ="Locations"
                            Caption ="Locations | Plate | Weight"
                            GridlineColor =0
                            LayoutCachedLeft =5265
                            LayoutCachedTop =475
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5349
                                    Top =475
                                    Width =22605
                                    Height =12960
                                    Name ="BulkChange_Locations"
                                    SourceObject ="Form.BulkChange_Locations"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5349
                                    LayoutCachedTop =475
                                    LayoutCachedWidth =27954
                                    LayoutCachedHeight =13435
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =475
                            Width =22755
                            Height =12965
                            Name ="COA"
                            Caption ="MSO | COA | POA | Title"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =475
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5346
                                    Top =475
                                    Width =22650
                                    Height =12960
                                    Name ="BulkChange_CoaMso"
                                    SourceObject ="Form.BulkChange_CoaMso"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5346
                                    LayoutCachedTop =475
                                    LayoutCachedWidth =27996
                                    LayoutCachedHeight =13435
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =480
                            Width =22755
                            Height =12960
                            Name ="Syndication"
                            Caption ="Interim Funding | Syndication"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =480
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5343
                                    Top =480
                                    Width =22590
                                    Height =12960
                                    Name ="BulkChange_Syndication"
                                    SourceObject ="Form.BulkChange_Syndication"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5343
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =27933
                                    LayoutCachedHeight =13440
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =479
                            Width =22755
                            Height =12961
                            BorderColor =0
                            Name ="ChildParts"
                            Caption ="Child Parts"
                            GridlineColor =0
                            LayoutCachedLeft =5265
                            LayoutCachedTop =479
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5349
                                    Top =479
                                    Width =22635
                                    Height =12960
                                    Name ="BulkChange_ChildParts"
                                    SourceObject ="Form.BulkChange_ChildParts"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5349
                                    LayoutCachedTop =479
                                    LayoutCachedWidth =27984
                                    LayoutCachedHeight =13439
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =479
                            Width =22755
                            Height =12961
                            BorderColor =0
                            Name ="InsertChildParts"
                            Caption ="Manage Child Parts"
                            GridlineColor =0
                            LayoutCachedLeft =5265
                            LayoutCachedTop =479
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5349
                                    Top =479
                                    Width =22605
                                    Height =12960
                                    Name ="BulkChange_ChildPartControl"
                                    SourceObject ="Form.BulkChange_ChildPartControl"
                                    VerticalAnchor =2

                                    LayoutCachedLeft =5349
                                    LayoutCachedTop =479
                                    LayoutCachedWidth =27954
                                    LayoutCachedHeight =13439
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5265
                            Top =475
                            Width =22755
                            Height =12965
                            Name ="TitleCheckin"
                            Caption ="Title Check In-Out"
                            LayoutCachedLeft =5265
                            LayoutCachedTop =475
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13440
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5349
                                    Top =475
                                    Width =22425
                                    Height =12960
                                    Name ="MassMod_Titles"
                                    SourceObject ="Form.MassMod_Titles"

                                    LayoutCachedLeft =5349
                                    LayoutCachedTop =475
                                    LayoutCachedWidth =27774
                                    LayoutCachedHeight =13435
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =6630
                    Width =720
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    Name ="SearchAssetID"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =6630
                    LayoutCachedWidth =780
                    LayoutCachedHeight =6870
                End
                Begin Label
                    OverlapFlags =93
                    Left =1290
                    Top =6300
                    Width =2595
                    Height =285
                    ForeColor =-2147483615
                    Name ="GroupInstructions"
                    Caption ="▲ Select Sch | Group ▲"
                    LayoutCachedLeft =1290
                    LayoutCachedTop =6300
                    LayoutCachedWidth =3885
                    LayoutCachedHeight =6585
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =223
                    Left =2640
                    Top =6315
                    Width =2520
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="GetSelectedUnits"
                    Caption ="▼Selected (0) Units ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2640
                    LayoutCachedTop =6315
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =6570
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverColor =10319446
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =780
                    Top =6630
                    Width =1872
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    Name ="SearchUnitVin"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =780
                    LayoutCachedTop =6630
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =6870
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =1140
                    Top =6315
                    Width =1470
                    Height =255
                    FontSize =10
                    TabIndex =6
                    Name ="GetAll"
                    Caption ="▼All 1"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1140
                    LayoutCachedTop =6315
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =6570
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
                    HoverColor =5026082
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =2652
                    Top =6630
                    Width =720
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    Name ="SearchUnitNum"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =2652
                    LayoutCachedTop =6630
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =6870
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =420
                    Width =2718
                    Height =2775
                    FontSize =9
                    TabIndex =8
                    BorderColor =14211288
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName AS Client, Count(vw_SixKe"
                        "ys.AssetID) AS Units, vw_SixKeys.ClientGroupId FROM vw_SixKeys GROUP BY vw_SixKe"
                        "ys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.ClientGroupId ORDER BY vw_Si"
                        "xKeys.ClientShortName;"
                    ColumnWidths ="0;1872;576;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =420
                    LayoutCachedWidth =2793
                    LayoutCachedHeight =3195
                    BackThemeColorIndex =1
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =120
                    Top =3360
                    Width =2715
                    Height =1320
                    FontSize =9
                    TabIndex =9
                    BorderColor =14211288
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig From vw_SixKeys WHERE ClientId =3674 GROUP BY MLID, ML"
                        "No, MLOrig Order by MLOrig, MLNo;"
                    ColumnWidths ="0;1872;576"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =3360
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =4680
                    BackThemeColorIndex =1
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2907
                    Top =210
                    Width =2280
                    Height =3000
                    FontSize =9
                    TabIndex =10
                    BorderColor =14211288
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchID, Schedule, Count(AssetID) AS Units From vw_SixKeys WHERE MLID =292 "
                        "GROUP BY SchID, Schedule Order by CDbl(NumInString(Left(Schedule,3))) Desc, Sche"
                        "dule Desc;"
                    ColumnWidths ="0;1584;576"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2907
                    LayoutCachedTop =210
                    LayoutCachedWidth =5187
                    LayoutCachedHeight =3210
                    BackThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =105
                    Top =3180
                    Width =2259
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="Label22"
                    Caption ="MLA#                                  Orig"
                    LayoutCachedLeft =105
                    LayoutCachedTop =3180
                    LayoutCachedWidth =2364
                    LayoutCachedHeight =3435
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =95
                    Left =2913
                    Width =2274
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Schedules                  Units"
                    LayoutCachedLeft =2913
                    LayoutCachedWidth =5187
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2907
                    Top =3360
                    Width =2280
                    Height =1320
                    FontSize =9
                    TabIndex =11
                    BorderColor =14211288
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SGrpID, UnitGroup, Count(AssetID) AS Units From vw_SixKeys WHERE SchID =2"
                        "255 GROUP BY SGrpID, UnitGroup Order by UnitGroup;"
                    ColumnWidths ="0;1584;576"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2907
                    LayoutCachedTop =3360
                    LayoutCachedWidth =5187
                    LayoutCachedHeight =4680
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =2913
                            Top =3180
                            Width =2274
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label97"
                            Caption ="Group                         Units"
                            LayoutCachedLeft =2913
                            LayoutCachedTop =3180
                            LayoutCachedWidth =5187
                            LayoutCachedHeight =3420
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =3372
                    Top =6630
                    Width =1822
                    FontSize =9
                    TabIndex =12
                    BackColor =13434879
                    Name ="SearchUnitDesc"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =3372
                    LayoutCachedTop =6630
                    LayoutCachedWidth =5194
                    LayoutCachedHeight =6870
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =780
                    Top =6870
                    Width =1872
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =13
                    ForeColor =8210719
                    Name ="ToggleSortUnitVIN"
                    DefaultValue ="0"
                    Caption ="↓↑ VIN"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =780
                    LayoutCachedTop =6870
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =7110
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =60
                    Top =6870
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =14
                    ForeColor =8210719
                    Name ="ToggleSortAssetID"
                    DefaultValue ="0"
                    Caption ="AssetID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =6870
                    LayoutCachedWidth =780
                    LayoutCachedHeight =7110
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =2652
                    Top =6870
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
                    ForeColor =8210719
                    Name ="ToggleSortUnitNum"
                    DefaultValue ="0"
                    Caption ="Unit #"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2652
                    LayoutCachedTop =6870
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =7110
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3372
                    Top =6870
                    Width =1822
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =16
                    ForeColor =8210719
                    Name ="ToggleSortUnitDesc"
                    DefaultValue ="0"
                    Caption ="Desc"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3372
                    LayoutCachedTop =6870
                    LayoutCachedWidth =5194
                    LayoutCachedHeight =7110
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =75
                    Top =6315
                    Width =1005
                    Height =255
                    FontWeight =400
                    TabIndex =17
                    ForeColor =8355711
                    Name ="PickModeToggle"
                    DefaultValue ="0"
                    Caption ="Pick Off"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =6315
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =6570
                    ForeTint =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15060409
                    BackTint =40.0
                    OldBorderStyle =1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =8
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Tab
                    OverlapFlags =215
                    BackStyle =0
                    Left =45
                    Top =4740
                    Width =5205
                    Height =1575
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    Name ="TabCtl131"

                    LayoutCachedLeft =45
                    LayoutCachedTop =4740
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =6315
                    UseTheme =1
                    OldBorderStyle =0
                    PressedThemeColorIndex =2
                    PressedShade =50.0
                    HoverForeColor =-2147483617
                    PressedForeColor =-2147483617
                    ForeThemeColorIndex =2
                    ForeShade =50.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =5160
                            Width =5055
                            Height =1080
                            Name ="UnitDesc"
                            Caption ="Unit Desc"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5160
                            LayoutCachedWidth =5175
                            LayoutCachedHeight =6240
                            Begin
                                Begin ListBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =195
                                    Top =5220
                                    Width =4902
                                    Height =945
                                    FontSize =9
                                    BorderColor =14211288
                                    Name ="DescList"
                                    RowSourceType ="Table/Query"
                                    ColumnWidths ="4435;504;0"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="2024 Peterbilt Daycab Tractor (A)"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =195
                                    LayoutCachedTop =5220
                                    LayoutCachedWidth =5097
                                    LayoutCachedHeight =6165
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =120
                            Top =5160
                            Width =5057
                            Height =1075
                            Name ="Draws"
                            Caption ="Draws"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5160
                            LayoutCachedWidth =5177
                            LayoutCachedHeight =6235
                            Begin
                                Begin ListBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    Left =125
                                    Top =5240
                                    Width =4317
                                    Height =945
                                    FontSize =9
                                    BorderColor =14211288
                                    Name ="DrawDownList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT U.DrawDownID, D.DrawNumber, Nz(D.DrawName,'None') AS DrawName, Count(U.Un"
                                        "itID) AS Units, U.GroupRef FROM Units AS U LEFT JOIN DrawDowns AS D ON U.DrawDow"
                                        "nID = D.DrawDownID Where (((U.GroupRef) = 5625)) GROUP BY U.DrawDownID, D.DrawNu"
                                        "mber, Nz([d].[DrawName],'None'), U.GroupRef;"
                                    ColumnWidths ="0;432;2880;432"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="2016 Freightliner Sleeper Tractor (D)"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =125
                                    LayoutCachedTop =5240
                                    LayoutCachedWidth =4442
                                    LayoutCachedHeight =6185
                                    BackThemeColorIndex =4
                                    BackTint =20.0
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =4500
                                    Top =5460
                                    Width =600
                                    Height =600
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Editdraw"
                                    Caption ="Edit  Draw"
                                    OnClick ="[Event Procedure]"
                                    GroupTable =2

                                    LayoutCachedLeft =4500
                                    LayoutCachedTop =5460
                                    LayoutCachedWidth =5100
                                    LayoutCachedHeight =6060
                                    LayoutGroup =1
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackThemeColorIndex =2
                                    BackTint =60.0
                                    BorderThemeColorIndex =1
                                    HoverColor =48840
                                    PressedColor =48840
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    GroupTable =2
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-625
                                End
                            End
                        End
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
' See "BulkChange_Form.cls"
