Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6353
    DatasheetFontHeight =11
    ItemSuffix =114
    Left =9945
    Top =2235
    Right =16755
    Bottom =14700
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xb57d1860d28ce440
    End
    Caption ="Add or Edit an Equipment Type"
    DatasheetFontName ="Calibri"
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
            FontSize =11
            FontName ="Calibri"
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
            Height =2220
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =1470
                    Width =3345
                    Height =540
                    FontSize =18
                    Name ="Label1"
                    Caption ="Equipment Types"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =1470
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1530
                    Top =1905
                    Width =3225
                    Height =285
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Type within Class"
                    GroupTable =2
                    LayoutCachedLeft =1530
                    LayoutCachedTop =1905
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =2190
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Top =45
                    Height =2160
                    FontSize =8
                    Name ="ClassList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [ClassTypeRecap].ID, [ClassTypeRecap].grpClass AS Class FROM [ClassTypeRe"
                        "cap] GROUP BY [ClassTypeRecap].ID, [ClassTypeRecap].grpClass ORDER BY [ClassType"
                        "Recap].grpClass;"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedTop =45
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =2205
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =30
                    Top =1905
                    Width =1057
                    Height =285
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Class"
                    GroupTable =2
                    LayoutCachedLeft =30
                    LayoutCachedTop =1905
                    LayoutCachedWidth =1087
                    LayoutCachedHeight =2190
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    Left =4815
                    Top =1905
                    Width =1260
                    Height =285
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Abreviaton"
                    GroupTable =2
                    LayoutCachedLeft =4815
                    LayoutCachedTop =1905
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =2190
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6315
                    Top =1905
                    Width =0
                    Height =285
                    Name ="EmptyCell53"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6315
                    LayoutCachedTop =1905
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =2190
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =5445
                    Top =90
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =5445
                    LayoutCachedTop =90
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =318
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =2
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =4860
                    Top =90
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
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

                    LayoutCachedLeft =4860
                    LayoutCachedTop =90
                    LayoutCachedWidth =5415
                    LayoutCachedHeight =318
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =2
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =6255
                    Top =1905
                    Width =0
                    Height =285
                    Name ="EmptyCell72"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6255
                    LayoutCachedTop =1905
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =2190
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6195
                    Top =1905
                    Width =0
                    Height =285
                    Name ="EmptyCell75"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6195
                    LayoutCachedTop =1905
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =2190
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6135
                    Top =1905
                    Width =0
                    Height =285
                    Name ="EmptyCell78"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6135
                    LayoutCachedTop =1905
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =2190
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =1140
                    Top =1905
                    Width =330
                    Height =285
                    Name ="EmptyCell81"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedTop =1905
                    LayoutCachedWidth =1470
                    LayoutCachedHeight =2190
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =1470
                    Top =585
                    Width =315
                    Height =360
                    FontSize =14
                    Name ="Label2467"
                    Caption =""
                    FontName ="Wingdings"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1470
                    LayoutCachedTop =585
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =945
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =1470
                    Top =945
                    Width =315
                    Height =555
                    FontSize =14
                    Name ="Label2468"
                    Caption =""
                    FontName ="Wingdings"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1470
                    LayoutCachedTop =945
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =1500
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1785
                    Top =585
                    Width =3525
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label120"
                    Caption ="Select a Class (from the list)"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1785
                    LayoutCachedTop =585
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =945
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1785
                    Top =945
                    Width =3525
                    Height =555
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label84"
                    Caption ="Edit the Types as needed.\015\012To ADD a Type, just type in the last row "
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1785
                    LayoutCachedTop =945
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =1500
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =1470
                    Top =1500
                    Width =315
                    Height =360
                    FontSize =14
                    ForeColor =1643706
                    Name ="Label2469"
                    Caption =""
                    FontName ="Wingdings"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1470
                    LayoutCachedTop =1500
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =1860
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1785
                    Top =1500
                    Width =3525
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =1643706
                    Name ="Label112"
                    Caption ="You May Delete Unused Types"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1785
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =1860
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =5
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =4800
                    Top =900
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="SaveUpdate"
                    Caption ="Save Update"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4800
                    LayoutCachedTop =900
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =1185
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
            End
        End
        Begin Section
            Visible = NotDefault
            CanGrow = NotDefault
            Height =345
            Name ="Detail"
            AlternateBackColor =16249583
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =30
                    Width =3225
                    Height =285
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    Name ="groupType"
                    ControlSource ="groupType"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1530
                    LayoutCachedTop =30
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4815
                    Top =30
                    Width =1260
                    Height =285
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    Name ="GrpTypAbrev"
                    ControlSource ="GrpTypAbrev"
                    GroupTable =2

                    LayoutCachedLeft =4815
                    LayoutCachedTop =30
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6135
                    Top =30
                    Width =0
                    Height =285
                    FontSize =8
                    TabIndex =3
                    Name ="ClassId"
                    ControlSource ="ClassId"
                    StatusBarText ="CLASS ASSOCIATED WITH UNIT TYPE FOR REPORTING"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Class is required"
                    GroupTable =2

                    LayoutCachedLeft =6135
                    LayoutCachedTop =30
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =30
                    Top =30
                    Width =1057
                    Height =285
                    FontSize =8
                    Name ="DeletedMe"
                    Caption ="Delete Me"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =1087
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =1643706
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6255
                    Top =30
                    Width =0
                    Height =285
                    FontSize =8
                    TabIndex =5
                    Name ="GroupClass"
                    ControlSource ="GrpClass"
                    StatusBarText ="CLASS ASSOCIATED WITH UNIT TYPE FOR REPORTING"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Class is required"
                    GroupTable =2

                    LayoutCachedLeft =6255
                    LayoutCachedTop =30
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6195
                    Top =30
                    Width =0
                    Height =285
                    FontSize =8
                    TabIndex =4
                    Name ="GroupTypeID"
                    ControlSource ="GroupTypeID"
                    StatusBarText ="CLASS ASSOCIATED WITH UNIT TYPE FOR REPORTING"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Class is required"
                    GroupTable =2

                    LayoutCachedLeft =6195
                    LayoutCachedTop =30
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6315
                    Top =30
                    Width =0
                    Height =285
                    Name ="EmptyCell80"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6315
                    LayoutCachedTop =30
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =1140
                    Top =30
                    Width =330
                    Height =285
                    Name ="EmptyCell82"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedTop =30
                    LayoutCachedWidth =1470
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
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
' See "Equipment Type_Add.cls"
