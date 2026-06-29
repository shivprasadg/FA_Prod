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
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14460
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =8550
    Top =2475
    Right =23010
    Bottom =7020
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6482071357eae340
    End
    DatasheetFontName ="Calibri"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =480
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =4350
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label5"
                    Caption ="VIN Verify - MSO Check In"
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =2
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12585
                    Top =60
                    Width =750
                    Height =303
                    FontSize =8
                    FontWeight =700
                    ForeColor =16777215
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

                    LayoutCachedLeft =12585
                    LayoutCachedTop =60
                    LayoutCachedWidth =13335
                    LayoutCachedHeight =363
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderColor =16777215
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    QuickStyle =22
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =13380
                    Top =60
                    Width =750
                    Height =303
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =16777215
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13380
                    LayoutCachedTop =60
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =363
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderColor =16777215
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    QuickStyle =22
                    QuickStyleMask =-629
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =4080
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =93
                    Left =2700
                    Top =30
                    TabIndex =3
                    Name ="LessThan17"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =30
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2925
                            Width =2235
                            Height =255
                            Name ="Label22"
                            Caption ="Less Than 17 Characters"
                            LayoutCachedLeft =2925
                            LayoutCachedWidth =5160
                            LayoutCachedHeight =255
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    TextFontFamily =50
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =990
                    Width =5100
                    Height =2925
                    FontSize =18
                    TabIndex =1
                    Name ="VIN_List"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;5040"
                    FontName ="OCR A Extended"

                    LayoutCachedLeft =60
                    LayoutCachedTop =990
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =3915
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =90
                            Top =675
                            Width =1575
                            Height =315
                            Name ="Label19"
                            Caption ="Matched VIN(s):"
                            LayoutCachedLeft =90
                            LayoutCachedTop =675
                            LayoutCachedWidth =1665
                            LayoutCachedHeight =990
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =5205
                    Top =270
                    Width =9135
                    Height =3645
                    TabIndex =2
                    Name ="FindVIN subform"
                    SourceObject ="Form.FindVIN subform"
                    LinkChildFields ="UnitID"
                    LinkMasterFields ="VIN_List"
                    EventProcPrefix ="FindVIN_subform"

                    LayoutCachedLeft =5205
                    LayoutCachedTop =270
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =3915
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontFamily =50
                    IMESentenceMode =3
                    Left =105
                    Top =255
                    Width =5055
                    Height =390
                    FontSize =18
                    BorderColor =-2147483615
                    Name ="seekvin"
                    FontName ="OCR A Extended"
                    OnChange ="[Event Procedure]"
                    ShowDatePicker =0

                    LayoutCachedLeft =105
                    LayoutCachedTop =255
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontFamily =0
                            Left =90
                            Top =15
                            Width =1815
                            Height =315
                            Name ="Label1"
                            Caption ="Search / Enter VIN:"
                            LayoutCachedLeft =90
                            LayoutCachedTop =15
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =330
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
' See "VINVerify.cls"
