Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3780
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =4770
    Top =4110
    Right =10335
    Bottom =8790
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xabee2f86970ae440
    End
    Caption ="Password Entry"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Password Entry"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    RibbonName ="Password Entry"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =795
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =780
                    Top =495
                    Width =2580
                    Height =300
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Password"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =780
                    LayoutCachedTop =495
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =795
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =720
                    Width =1005
                    Height =720
                    FontSize =24
                    FontWeight =700
                    Name ="Label466"
                    Caption ="OPS"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =720
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =720
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =1320
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =840
                    Top =120
                    Width =2640
                    Height =300
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Enter Password"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =840
                    LayoutCachedTop =120
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =480
                    Width =2640
                    Height =315
                    Name ="Password"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="Password"

                    LayoutCachedLeft =840
                    LayoutCachedTop =480
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =795
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =840
                    Top =900
                    Width =1260
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="SubmitRequest"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =840
                    LayoutCachedTop =900
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1200
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =9
                    QuickStyleMask =-129
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2220
                    Top =900
                    Width =1260
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="CancelRequest"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2220
                    LayoutCachedTop =900
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1200
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =9
                    QuickStyleMask =-129
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
' See "OpsPassword.cls"
