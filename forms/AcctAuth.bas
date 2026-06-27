Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5055
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =10350
    Top =2775
    Right =15405
    Bottom =4920
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x256dbd69d0e0e340
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin FormHeader
            Height =885
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Top =495
                    Width =5040
                    Height =375
                    BackColor =8210719
                    Name ="Box6"
                    LayoutCachedTop =495
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =870
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =3
                    Left =2250
                    Top =15
                    Width =2805
                    Height =375
                    FontSize =16
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label4"
                    Caption ="Accounting Security"
                    LayoutCachedLeft =2250
                    LayoutCachedTop =15
                    LayoutCachedWidth =5055
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =3915
                    Top =540
                    Width =1035
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3915
                    LayoutCachedTop =540
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =828
                    PictureCaptionArrangement =5
                    UseTheme =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =204
                    Left =315
                    Top =480
                    Width =3045
                    Height =345
                    FontSize =14
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label1"
                    Caption ="Enter Authorization Code"
                    LayoutCachedLeft =315
                    LayoutCachedTop =480
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =825
                End
            End
        End
        Begin Section
            Height =1275
            Name ="Detail"
            Begin
                Begin Image
                    PictureType =2
                    Left =90
                    Width =1545
                    Height =1275
                    Name ="Image7"
                    Picture ="ATLAASGlobeLogo"

                    LayoutCachedLeft =90
                    LayoutCachedWidth =1635
                    LayoutCachedHeight =1275
                    TabIndex =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1710
                    Top =165
                    Width =3075
                    Height =330
                    FontSize =14
                    FontWeight =700
                    BackColor =13434879
                    BorderColor =62207
                    ForeColor =-2147483617
                    Name ="AcctPW"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="Password"

                    LayoutCachedLeft =1710
                    LayoutCachedTop =165
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =495
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =1710
                    Top =630
                    Width =3060
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    ForeColor =1643706
                    Name ="AcctSubmitBtn"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =1710
                    LayoutCachedTop =630
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =885
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =1643706
                    HoverThemeColorIndex =1
                    HoverShade =95.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =1643706
                    PressedForeColor =1643706
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
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
' See "AcctAuth.cls"
