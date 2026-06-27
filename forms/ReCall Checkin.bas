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
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9030
    DatasheetFontHeight =11
    ItemSuffix =22
    Left =5760
    Top =2790
    Right =18435
    Bottom =10425
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =348
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =360
                    Width =3990
                    Height =345
                    FontSize =16
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label5"
                    Caption ="Recall Check In"
                    LayoutCachedLeft =360
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =345
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =8055
                    Top =60
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    ForeColor =16777215
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =8055
                    LayoutCachedTop =60
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =288
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7470
                    Top =60
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =16777215
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
                    BackStyle =0

                    LayoutCachedLeft =7470
                    LayoutCachedTop =60
                    LayoutCachedWidth =8025
                    LayoutCachedHeight =288
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =1185
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =7530
                    Top =60
                    Width =1290
                    Height =390
                    TabIndex =1
                    Name ="findvinbtn"
                    Caption ="Submit VIN"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =14
                        Begin
                            Action ="OpenForm"
                            Argument ="LogRecallf"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![logrecallf]![recallref]"
                            Argument ="[Forms]![recall checkin]![recallno]"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![logrecallf]![rcunitid]"
                            Argument ="[Forms]![recall checkin]![seekvin]"
                        End
                        Begin
                            Action ="Close"
                            Argument ="2"
                            Argument ="LogRecallf"
                            Argument ="0"
                        End
                        Begin
                            Action ="MsgBox"
                            Argument ="Unit has been added. Select another unit or close form"
                            Argument ="-1"
                            Argument ="4"
                            Argument ="Unit Added to Recall"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"findvinbtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">LogRecallf</A"
                        End
                        Begin
                            Comment ="_AXL:rgument><Argument Name=\"DataMode\">Add</Argument></Action><Action Name=\"S"
                                "etValue\"><Argument Name=\"Item\">[Forms]![logrecallf]![recallref]</Argument><Ar"
                                "gument Name=\"Expression\">[Forms]![recall checkin]![recallno]</Argument></Actio"
                                "n><Action Name=\"SetVa"
                        End
                        Begin
                            Comment ="_AXL:lue\"><Argument Name=\"Item\">[Forms]![logrecallf]![rcunitid]</Argument><Ar"
                                "gument Name=\"Expression\">[Forms]![recall checkin]![seekvin]</Argument></Action"
                                "><Action Name=\"CloseWindow\"><Argument Name=\"ObjectType\">Form</Argument><Argu"
                                "ment Name=\"ObjectName\""
                        End
                        Begin
                            Comment ="_AXL:>LogRecallf</Argument></Action><Action Name=\"MessageBox\"><Argument Name=\""
                                "Message\">Unit has been added. Select another unit or close form</Argument><Argu"
                                "ment Name=\"Type\">Information</Argument><Argument Name=\"Title\">Unit Added to "
                                "Recall</Argument></A"
                        End
                        Begin
                            Comment ="_AXL:ction></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7530
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =7920
                    Left =1575
                    Top =525
                    Width =4185
                    Height =315
                    TabIndex =2
                    Name ="recallno"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RecallCampaign.ID, RecallCampaign.RecallNo, RecallCampaign.RecallDesc, Re"
                        "callCampaign.RecallDt FROM RecallCampaign ORDER BY RecallCampaign.RecallDt DESC "
                        ", RecallCampaign.RecallNo;"
                    ColumnWidths ="0;2160;2880;2880"
                    ValidationRule =">0"
                    ValidationText ="Recall or other Campaign must be selected"
                    Format ="General Number"

                    LayoutCachedLeft =1575
                    LayoutCachedTop =525
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =45
                            Top =510
                            Width =3060
                            Height =315
                            Name ="Label19"
                            Caption ="Select Recall:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =510
                            LayoutCachedWidth =3105
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1575
                    Top =30
                    Width =4185
                    Height =405
                    FontSize =14
                    BorderColor =-2147483615
                    Name ="seekvin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Units.UnitID, Units.UnitVIN FROM Units WHERE (((Units.UnitVIN) Is Not Nul"
                        "l) AND ((Units.UnitStatus)<>\"k\" And (Units.UnitStatus)<>\"ks\"));"
                    ColumnWidths ="0;2880"

                    LayoutCachedLeft =1575
                    LayoutCachedTop =30
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =435
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =90
                            Top =30
                            Width =3060
                            Height =405
                            FontSize =12
                            Name ="Label1"
                            Caption ="Enter VIN:"
                            LayoutCachedLeft =90
                            LayoutCachedTop =30
                            LayoutCachedWidth =3150
                            LayoutCachedHeight =435
                        End
                    End
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
' See "ReCall Checkin.cls"
