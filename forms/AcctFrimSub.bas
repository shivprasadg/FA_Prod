Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4395
    DatasheetFontHeight =11
    ItemSuffix =5
    Left =8430
    Top =12900
    Right =13005
    Bottom =14175
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x3976ed19d407e440
    End
    RecordSource ="SELECT ClientAcctFirm.CAcctClientRef, ClientAcctFirm.CAAcct, ClientAcctFirm.AFDT"
        "ADD FROM ClientAcctFirm;"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin FormHeader
            Height =390
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =2475
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483609
                    Name ="Label13"
                    Caption ="Accounting Firm(s) Name"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =300
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2685
                    Top =45
                    Width =1170
                    Height =285
                    FontSize =8
                    FontWeight =700
                    Name ="AddAcctFBtn"
                    Caption ="Add  New"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =14
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![AcctFirmSub]![CAcctClientRef]=[Forms]![Clients_view]![clientID]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddAcctFBtn\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"GoToRecord\"><Argu"
                                "ment Name=\"Record\">New</Argumen"
                        End
                        Begin
                            Comment ="_AXL:t></Action><Action Name=\"SetValue\"><Argument Name=\"Item\">[Forms]![AcctF"
                                "irmSub]![CAcctClientRef]=[Forms]![Clients_view]![clientID]</Argument></Action></"
                                "Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2685
                    LayoutCachedTop =45
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    ThemeFontIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-65
                End
            End
        End
        Begin Section
            Height =555
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =90
                    Width =3210
                    Height =315
                    FontSize =10
                    FontWeight =700
                    Name ="CAAcct"
                    ControlSource ="CAAcct"
                    RowSourceType ="Table/Query"
                    RowSource ="AcctFirmNm"
                    StatusBarText ="accounting firm"
                    ListItemsEditForm ="AcctNMForm"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =90
                    LayoutCachedWidth =4230
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =225
                            Top =90
                            Width =780
                            Height =315
                            FontSize =9
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label1"
                            Caption ="FIRM:"
                            LayoutCachedLeft =225
                            LayoutCachedTop =90
                            LayoutCachedWidth =1005
                            LayoutCachedHeight =405
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =195
                    Top =360
                    Width =390
                    Height =135
                    TabIndex =1
                    Name ="CAcctClientRef"
                    ControlSource ="CAcctClientRef"
                    StatusBarText ="cleint id"

                    LayoutCachedLeft =195
                    LayoutCachedTop =360
                    LayoutCachedWidth =585
                    LayoutCachedHeight =495
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
