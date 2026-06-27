Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =1
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12780
    DatasheetFontHeight =11
    ItemSuffix =360
    Left =7050
    Top =3705
    Right =22710
    Bottom =19815
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x828b3cb36f56e640
    End
    RecordSource ="SELECT WorkOrders.wAssetID, * FROM WorkOrders WHERE (((WorkOrders.wAssetID)>0));"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
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
                    Width =10598
                    Height =210
                    BorderColor =16777215
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =360
                    LayoutCachedWidth =10598
                    LayoutCachedHeight =570
                    TabIndex =5
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =120
                    Width =1740
                    Height =420
                    FontSize =16
                    Name ="Label1053"
                    Caption ="Work Order"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =119
                    Left =8400
                    Top =60
                    Width =1560
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="AddNew"
                    Caption ="Add a Work Order"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6480
                    Top =60
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="WorkOrderID"
                    ControlSource ="WorkOrderID"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="wAssetId"
                    ControlSource ="wAssetId"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =4860
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5640
                    Top =60
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="wClientID"
                    ControlSource ="wClientID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =5640
                    LayoutCachedTop =60
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7260
                    Top =60
                    Width =780
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="wClientIDGroupID"
                    ControlSource ="wClientIDGroupID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5580
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =780
                    Width =12660
                    Height =4665
                    BorderColor =16777215
                    Name ="WorkOrderItemsSub"
                    SourceObject ="Form.WorkOrderItemsSub"
                    LinkChildFields ="WorkOrderID"
                    LinkMasterFields ="WorkOrderID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =780
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =5445
                    BorderShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4443
                    Height =330
                    FontSize =12
                    TabIndex =2
                    Name ="wStartDate"
                    ControlSource ="wStartDate"
                    Format ="Short Date"
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
                                "nterfaceMacro For=\"wStartDate\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">AddRmktExpFor"
                        End
                        Begin
                            Comment ="_AXL:m</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4443
                    LayoutCachedWidth =5883
                    LayoutCachedHeight =330
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =3948
                            Width =495
                            Height =330
                            FontSize =10
                            Name ="Label108"
                            Caption ="Start:"
                            LayoutCachedLeft =3948
                            LayoutCachedWidth =4443
                            LayoutCachedHeight =330
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Width =2748
                    Height =330
                    FontSize =12
                    TabIndex =1
                    Name ="wONumber"
                    ControlSource ="wONumber"

                    LayoutCachedLeft =1200
                    LayoutCachedWidth =3948
                    LayoutCachedHeight =330
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
                            Left =60
                            Width =1080
                            Height =330
                            Name ="Label4"
                            Caption ="Order Num "
                            LayoutCachedLeft =60
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =330
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =660
                    Top =390
                    Width =9720
                    Height =330
                    FontSize =10
                    TabIndex =5
                    Name ="wNotes"
                    ControlSource ="wNotes"

                    LayoutCachedLeft =660
                    LayoutCachedTop =390
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =390
                            Width =540
                            Height =330
                            FontSize =12
                            Name ="Label152"
                            Caption ="Notes "
                            LayoutCachedLeft =60
                            LayoutCachedTop =390
                            LayoutCachedWidth =600
                            LayoutCachedHeight =720
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
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6423
                    Height =330
                    FontSize =12
                    TabIndex =4
                    Name ="wCloseDate"
                    ControlSource ="wCloseDate"
                    Format ="Short Date"
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
                                "nterfaceMacro For=\"wCloseDate\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">AddRmktExpFor"
                        End
                        Begin
                            Comment ="_AXL:m</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =6423
                    LayoutCachedWidth =7863
                    LayoutCachedHeight =330
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =5883
                            Width =540
                            Height =330
                            FontSize =10
                            Name ="Label285"
                            Caption ="Close"
                            LayoutCachedLeft =5883
                            LayoutCachedWidth =6423
                            LayoutCachedHeight =330
                            ColumnStart =4
                            ColumnEnd =4
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8478
                    Width =1485
                    Height =330
                    FontSize =12
                    TabIndex =3
                    Name ="wStatus"
                    ControlSource ="wStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"Pending\";\"Complete\""
                    StatusBarText ="unit link"

                    LayoutCachedLeft =8478
                    LayoutCachedWidth =9963
                    LayoutCachedHeight =330
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =7863
                            Width =615
                            Height =330
                            FontSize =12
                            Name ="Label283"
                            Caption ="Status"
                            LayoutCachedLeft =7863
                            LayoutCachedWidth =8478
                            LayoutCachedHeight =330
                            ColumnStart =6
                            ColumnEnd =6
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
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
' See "WorkOrdersForm.cls"
