Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14910
    DatasheetFontHeight =11
    ItemSuffix =61
    Left =11010
    Top =14865
    Right =26040
    Bottom =17370
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x78c77746b021e540
    End
    RecordSource ="SELECT mlamends.ID, mlamends.MLRef, mlamends.MLAmtNO, mlamends.MLAmdDt, mlamends"
        ".MLAmdDesc FROM mlamends ORDER BY mlamends.MLAmdDt DESC;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =15707392
            Name ="FormHeader"
        End
        Begin Section
            Height =2538
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3672
                    Top =600
                    Width =960
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="MLAmtNO"
                    ControlSource ="MLAmtNO"
                    StatusBarText ="amendment number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =3672
                    LayoutCachedTop =600
                    LayoutCachedWidth =4632
                    LayoutCachedHeight =885
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2280
                            Top =600
                            Width =1320
                            Height =285
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label1"
                            Caption ="Amendment No."
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =2280
                            LayoutCachedTop =600
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =885
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5244
                    Top =600
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="MLAmdDt"
                    ControlSource ="MLAmdDt"
                    Format ="Short Date"
                    StatusBarText ="amendment date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =5244
                    LayoutCachedTop =600
                    LayoutCachedWidth =6684
                    LayoutCachedHeight =885
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4692
                            Top =600
                            Width =480
                            Height =285
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label2"
                            Caption =" Date"
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =4692
                            LayoutCachedTop =600
                            LayoutCachedWidth =5172
                            LayoutCachedHeight =885
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3672
                    Top =1080
                    Width =10788
                    Height =1308
                    FontSize =10
                    TabIndex =3
                    Name ="MLAmdDesc"
                    ControlSource ="MLAmdDesc"
                    StatusBarText ="descripiton of the amendment"
                    FontName ="Segoe UI"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =3672
                    LayoutCachedTop =1080
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =2388
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2280
                            Top =1080
                            Width =1320
                            Height =1308
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label3"
                            Caption =" Description"
                            FontName ="Segoe UI"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =2280
                            LayoutCachedTop =1080
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =2388
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8670
                    Top =345
                    Width =870
                    Height =315
                    TabIndex =5
                    Name ="MLRef"
                    ControlSource ="MLRef"
                    StatusBarText ="ML reference id"

                    LayoutCachedLeft =8670
                    LayoutCachedTop =345
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =345
                    Width =705
                    Height =315
                    TabIndex =6
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =345
                    LayoutCachedWidth =10305
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                End
                Begin EmptyCell
                    Left =6744
                    Top =600
                    Height =285
                    Name ="EmptyCell30"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =6744
                    LayoutCachedTop =600
                    LayoutCachedWidth =8184
                    LayoutCachedHeight =885
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8244
                    Top =600
                    Height =285
                    Name ="EmptyCell32"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =8244
                    LayoutCachedTop =600
                    LayoutCachedWidth =9684
                    LayoutCachedHeight =885
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =9744
                    Top =600
                    Width =4716
                    Height =285
                    Name ="EmptyCell34"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =9744
                    LayoutCachedTop =600
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =885
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =90
                    Top =360
                    Width =14760
                    Height =2115
                    Name ="Box36"
                    LayoutCachedLeft =90
                    LayoutCachedTop =360
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =2475
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =210
                    Top =120
                    Width =1500
                    Height =360
                    FontSize =12
                    Name ="Label0"
                    Caption ="Amendments"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedTop =120
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =163
                    Left =12060
                    Top =540
                    Width =2370
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =-2147483615
                    Name ="MLAmdBtn"
                    Caption ="Add New Amendment"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"MLAmdBtn\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Action Na"
                                "me=\"GoToRecord\"><Argument Name="
                        End
                        Begin
                            Comment ="_AXL:\"Record\">New</Argument></Action><ConditionalBlock><If><Condition>[MacroEr"
                                "ror]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argument Name="
                                "\"Message\">=[MacroError].[Description]</Argument></Action></Statements></If></C"
                                "onditionalBlock></S"
                        End
                        Begin
                            Comment ="_AXL:tatements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =12060
                    LayoutCachedTop =540
                    LayoutCachedWidth =14430
                    LayoutCachedHeight =930
                    UseTheme =1
                    HoverColor =10319446
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =360
                    Top =600
                    Width =1860
                    Height =1788
                    FontSize =10
                    TabIndex =2
                    Name ="AmendmentsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT mlamends.ID, mlamends.MLAmtNO AS AmendNum, mlamends.MLAmdDt AS AmendDate,"
                        " mlamends.MLRef From mlamends WHERE MLRef =6 ORDER BY MLAmtNO DESC , MLAmdDt DES"
                        "C;"
                    ColumnWidths ="0;576;864;0"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =360
                    LayoutCachedTop =600
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2388
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
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
' See "MLAmendf.cls"
