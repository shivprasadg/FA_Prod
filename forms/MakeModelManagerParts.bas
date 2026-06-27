Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15750
    DatasheetFontHeight =11
    ItemSuffix =223
    Left =5115
    Top =4215
    Right =20580
    Bottom =13335
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x81410d513732e540
    End
    RecordSource ="SELECT LookUp_ChildPartType.ChildPartType, * FROM LookUp_ChildPartType;"
    Caption ="Child Part Manager"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =255
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
            Height =4845
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =12720
                    Top =780
                    Width =840
                    Height =3120
                    BorderColor =8872517
                    Name ="AddBox"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =780
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =3900
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =120
                    Width =3525
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label1"
                    Caption ="Parts Management"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3645
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =75
                    Top =4230
                    Width =300
                    Height =510
                    FontSize =10
                    Name ="Label2468"
                    Caption ="❷"
                    FontName ="Segoe UI"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =75
                    LayoutCachedTop =4230
                    LayoutCachedWidth =375
                    LayoutCachedHeight =4740
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =75
                    Top =435
                    Width =315
                    Height =375
                    FontSize =10
                    Name ="Label2467"
                    Caption ="❶"
                    FontName ="Segoe UI"
                    GroupTable =16
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =75
                    LayoutCachedTop =435
                    LayoutCachedWidth =390
                    LayoutCachedHeight =810
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =16
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5985
                    Top =4230
                    Width =2280
                    Height =510
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =1643706
                    Name ="Label112"
                    Caption ="You May Delete Unused Models"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5985
                    LayoutCachedTop =4230
                    LayoutCachedWidth =8265
                    LayoutCachedHeight =4740
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5685
                    Top =4230
                    Width =300
                    Height =510
                    FontSize =10
                    ForeColor =2366701
                    Name ="Label2469"
                    Caption ="❸"
                    FontName ="Segoe UI"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5685
                    LayoutCachedTop =4230
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =4740
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =375
                    Top =4230
                    Width =5310
                    Height =510
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label84"
                    Caption ="Edit the Models on the sub form below, as needed.\015\012To ADD a Model, just ty"
                        "pe in the NEW model into the last row "
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =375
                    LayoutCachedTop =4230
                    LayoutCachedWidth =5685
                    LayoutCachedHeight =4740
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =12
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =390
                    Top =435
                    Width =2550
                    Height =375
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label120"
                    Caption ="Select a Make (from the list)"
                    GroupTable =16
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =390
                    LayoutCachedTop =435
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =810
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =16
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =13020
                    Top =1500
                    Width =1680
                    Height =1680
                    FontSize =10
                    TabIndex =3
                    Name ="AddMissingModel"
                    Caption ="Add Model to Make"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13020
                    LayoutCachedTop =1500
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =3180
                    PictureCaptionArrangement =3
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedColor =8210719
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3240
                    Top =810
                    Width =3135
                    Height =3420
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =2
                    Name ="ModelList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2880"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =810
                    LayoutCachedWidth =6375
                    LayoutCachedHeight =4230
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =9600
                    Top =780
                    Width =3240
                    Height =3780
                    ColumnOrder =4
                    FontSize =10
                    ForeColor =5855577
                    BorderColor =4144959
                    Name ="NoMatchModelsList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1728;576;576;0"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =780
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =4560
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeTint =65.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =6975
                    Top =795
                    Width =300
                    Height =870
                    FontSize =10
                    Name ="Label155"
                    Caption ="❹"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6975
                    LayoutCachedTop =795
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =1665
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =7275
                    Top =795
                    Width =2280
                    Height =870
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label156"
                    Caption ="You can add in Models from exisiting units from HERE →"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7275
                    LayoutCachedTop =795
                    LayoutCachedWidth =9555
                    LayoutCachedHeight =1665
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13260
                    Top =3960
                    Width =2340
                    Height =225
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483609
                    Name ="Make"
                    BottomPadding =38

                    LayoutCachedLeft =13260
                    LayoutCachedTop =3960
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =4185
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13260
                    Top =4305
                    Width =2340
                    Height =225
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =5
                    BorderColor =-2147483609
                    Name ="Model"
                    BottomPadding =38

                    LayoutCachedLeft =13260
                    LayoutCachedTop =4305
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =4530
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    Left =9600
                    Top =240
                    Width =3240
                    Height =540
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label179"
                    Caption ="Models in use. \015\012(but not part of Model's List)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9600
                    LayoutCachedTop =240
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =810
                    Width =2883
                    Height =3420
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =1
                    BackColor =15921906
                    Name ="MakeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartMake.ChildPartMakeID, LookUp_ChildPartMake.ChildPartMake "
                        "FROM LookUp_ChildPartMake ORDER BY LookUp_ChildPartMake.ChildPartMake;"
                    ColumnWidths ="0;2448"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =810
                    LayoutCachedWidth =2958
                    LayoutCachedHeight =4230
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =13740
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
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

                    LayoutCachedLeft =13740
                    LayoutCachedTop =120
                    LayoutCachedWidth =14295
                    LayoutCachedHeight =348
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
                    Left =14307
                    Top =120
                    Width =555
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
                    HorizontalAnchor =1

                    LayoutCachedLeft =14307
                    LayoutCachedTop =120
                    LayoutCachedWidth =14862
                    LayoutCachedHeight =348
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
            End
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =7935
            Name ="Detail"
            AlternateBackColor =16249583
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =120
                    Top =60
                    Width =15630
                    Height =7875
                    Name ="TabControl"
                    FontName ="Calibri"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =15750
                    LayoutCachedHeight =7935
                    ThemeFontIndex =1
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =195
                            Top =540
                            Width =15480
                            Height =7320
                            Name ="Models"
                            LayoutCachedLeft =195
                            LayoutCachedTop =540
                            LayoutCachedWidth =15675
                            LayoutCachedHeight =7860
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =240
                                    Top =660
                                    Width =15180
                                    Height =6600
                                    Name ="UnitMake"
                                    SourceObject ="Form.MakeModelManagerParts_Models"
                                    LinkChildFields ="ChildPartMakeID"
                                    LinkMasterFields ="MakeList"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =660
                                    LayoutCachedWidth =15420
                                    LayoutCachedHeight =7260
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =540
                            Width =15480
                            Height =7320
                            Name ="Makes"
                            LayoutCachedLeft =195
                            LayoutCachedTop =540
                            LayoutCachedWidth =15675
                            LayoutCachedHeight =7860
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =420
                                    Top =810
                                    Width =12150
                                    Height =6990
                                    Name ="UnitMakes"
                                    SourceObject ="Form.MakeModelManagerParts_Makes"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =810
                                    LayoutCachedWidth =12570
                                    LayoutCachedHeight =7800
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =540
                            Width =15483
                            Height =7320
                            Name ="UnitsUsing"
                            Caption ="Units Using Selected Child Model"
                            LayoutCachedLeft =195
                            LayoutCachedTop =540
                            LayoutCachedWidth =15678
                            LayoutCachedHeight =7860
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =270
                                    Top =540
                                    Width =15408
                                    Height =7260
                                    Name ="UnitPartsList_SubForm"
                                    SourceObject ="Form.UnitPartsList_SubForm"
                                    LinkChildFields ="ChildMake;ChildClass"
                                    LinkMasterFields ="Make;Model"

                                    LayoutCachedLeft =270
                                    LayoutCachedTop =540
                                    LayoutCachedWidth =15678
                                    LayoutCachedHeight =7800
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "MakeModelManagerParts.cls"
