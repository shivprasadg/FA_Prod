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
    ViewsAllowed =1
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =7710
    Top =3930
    Right =24240
    Bottom =14670
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x4e2eee0e136be440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
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
        Begin FormHeader
            Height =525
            BackColor =8210719
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =135
                    Width =4110
                    Height =525
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label5"
                    Caption ="C-of-A Check-in"
                    LayoutCachedLeft =135
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =525
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =420
                    Width =14400
                    Height =105
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="1_DkBlue_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =420
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =525
                    TabIndex =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12810
                    Top =75
                    Width =555
                    Height =228
                    FontSize =8
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

                    LayoutCachedLeft =12810
                    LayoutCachedTop =75
                    LayoutCachedWidth =13365
                    LayoutCachedHeight =303
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
                    Left =13381
                    Top =75
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13381
                    LayoutCachedTop =75
                    LayoutCachedWidth =14131
                    LayoutCachedHeight =303
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
        Begin Section
            CanGrow = NotDefault
            Height =9960
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =900
                    Width =2981
                    Height =5130
                    Name ="MasterLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.MLID, vw_SixKeys.MLNo FROM vw_SixKeys WHERE (((vw_SixKeys.Unit"
                        "ID) Is Not Null) AND ((vw_SixKeys.[MLOrig])=\"FA\") AND ((vw_SixKeys.status)=\"A"
                        "\" Or (vw_SixKeys.status)=\"B\")) GROUP BY vw_SixKeys.MLID, vw_SixKeys.MLNo, vw_"
                        "SixKeys.MLOrig HAVING (((vw_SixKeys.MLNo) Not Like \"demo*\")) ORDER BY vw_SixKe"
                        "ys.MLNo;"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =900
                    LayoutCachedWidth =3056
                    LayoutCachedHeight =6030
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =75
                            Top =660
                            Width =1335
                            Height =300
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Master Lease"
                            LayoutCachedLeft =75
                            LayoutCachedTop =660
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =75
                    Top =6330
                    Width =2981
                    Height =3450
                    TabIndex =1
                    Name ="SchID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo FROM MstrLease INNER JOIN Schedule ON Mstr"
                        "Lease.MLID = Schedule.SchMLRef WHERE (((MstrLease.MLID)=[forms]![cofa checkin]!["
                        "ml])) ORDER BY Schedule.SchID DESC;"
                    ColumnWidths ="0;0;1440"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =6330
                    LayoutCachedWidth =3056
                    LayoutCachedHeight =9780
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =75
                            Top =6060
                            Width =1335
                            Height =360
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Schedule"
                            LayoutCachedLeft =75
                            LayoutCachedTop =6060
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =6420
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =3240
                    Top =495
                    Width =11088
                    Height =9300
                    TabIndex =2
                    Name ="Unit_List_CofASub"
                    SourceObject ="Form.Unit List_CofASub"
                    LinkChildFields ="SGrpID"
                    LinkMasterFields ="SchID"
                    VerticalAnchor =2

                    LayoutCachedLeft =3240
                    LayoutCachedTop =495
                    LayoutCachedWidth =14328
                    LayoutCachedHeight =9795
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =75
                    Top =240
                    Width =2265
                    Height =270
                    TabIndex =3
                    BorderColor =8210719
                    Name ="SearchLease"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =240
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =75
                            Width =2100
                            Height =225
                            FontSize =8
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label29"
                            Caption ="Type SEARCH ▼ HERE"
                            LayoutCachedLeft =75
                            LayoutCachedWidth =2175
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =3255
                    Top =60
                    Width =3480
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="HeaderLabel"
                    Caption ="Unit's VIN Number List"
                    LayoutCachedLeft =3255
                    LayoutCachedTop =60
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =510
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =3300
                    Top =960
                    Width =2970
                    Height =525
                    FontSize =18
                    FontWeight =700
                    ForeColor =-2147483611
                    Name ="InstructionLabel1"
                    Caption ="◄ Select a Lease\015\012 "
                    LayoutCachedLeft =3300
                    LayoutCachedTop =960
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =1485
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =3300
                    Top =6645
                    Width =3690
                    Height =465
                    FontSize =18
                    FontWeight =700
                    ForeColor =-2147483611
                    Name ="InstructionLabel2"
                    Caption ="◄Select a Schedule"
                    LayoutCachedLeft =3300
                    LayoutCachedTop =6645
                    LayoutCachedWidth =6990
                    LayoutCachedHeight =7110
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "CofA Checkin.cls"
