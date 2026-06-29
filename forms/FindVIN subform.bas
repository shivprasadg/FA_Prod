Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8490
    DatasheetFontHeight =11
    ItemSuffix =120
    Left =14145
    Top =4590
    Right =23025
    Bottom =7980
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x45987a0616ade540
    End
    RecordSource ="SELECT Units.UnitID, Units.UnitVIN, Units.UnitYr, Units.UnitVINLocked, Units.Uni"
        "tStatus, Units.UnitVINVer, Units.UnitMSORecd, Units.UnitMake, Units.UnitModel, U"
        "nits.UnitMSOCopy, Units.UnitGVW, Locations.LState AS TitledState FROM Units LEFT"
        " JOIN Locations ON Units.UnitTitleLocationRef = Locations.LocationID WHERE (((Un"
        "its.UnitStatus)=\"B\" Or (Units.UnitStatus)=\"A\" Or (Units.UnitStatus)=\"TR\"))"
        ";"
    Caption ="FindVIN subform"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =204
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =465
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =15
                    Width =5625
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="Label9"
                    Caption ="Verified Update"
                    LayoutCachedLeft =60
                    LayoutCachedTop =15
                    LayoutCachedWidth =5685
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =2
                    ForeShade =50.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6135
                    Top =15
                    Width =1065
                    Height =375
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="ClosVvsub"
                    Caption ="CLOSE"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =14
                        Begin
                            Condition ="[UnitMSORecd] Is Null And [unitmsocopy] Is Null"
                            Action ="MsgBox"
                            Argument ="MSO or MSO received date is required"
                            Argument ="-1"
                            Argument ="1"
                            Argument ="Vin Not Verified"
                        End
                        Begin
                            Condition ="..."
                            Action ="Close"
                            Argument ="2"
                            Argument ="FindVIN subform"
                            Argument ="0"
                        End
                        Begin
                            Condition ="..."
                            Action ="StopMacro"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[unitvinver]"
                            Argument ="Date()"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[UnitVINLocked]"
                            Argument ="-1"
                        End
                        Begin
                            Action ="Close"
                            Argument ="2"
                            Argument ="FindVIN subform"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![VINVerify]![seekvin]"
                            Argument ="Null"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![VINVerify]![MY]"
                            Argument ="Null"
                        End
                        Begin
                            Action ="SetValue"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"ClosVvsub\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition>[U"
                                "nitMSORecd] Is Null And [unit"
                        End
                        Begin
                            Comment ="_AXL:msocopy] Is Null</Condition><Statements><Action Name=\"MessageBox\"><Argume"
                                "nt Name=\"Message\">MSO or MSO received date is required</Argument><Argument Nam"
                                "e=\"Type\">Critical</Argument><Argument Name=\"Title\">Vin Not Verified</Argumen"
                                "t></Action><Action Na"
                        End
                        Begin
                            Comment ="_AXL:me=\"CloseWindow\"><Argument Name=\"ObjectType\">Form</Argument><Argument N"
                                "ame=\"ObjectName\">FindVIN subform</Argument></Action><Action Name=\"StopMacro\""
                                "/></Statements></If></ConditionalBlock><Action Name=\"SetValue\"><Argument Name="
                                "\"Item\">[unitvinver]</Ar"
                        End
                        Begin
                            Comment ="_AXL:gument><Argument Name=\"Expression\">Date()</Argument></Action><Action Name"
                                "=\"SetValue\"><Argument Name=\"Item\">[UnitVINLocked]</Argument><Argument Name=\""
                                "Expression\">-1</Argument></Action><Action Name=\"CloseWindow\"><Argument Name=\""
                                "ObjectType\">Form</Argu"
                        End
                        Begin
                            Comment ="_AXL:ment><Argument Name=\"ObjectName\">FindVIN subform</Argument></Action><Acti"
                                "on Name=\"SetValue\"><Argument Name=\"Item\">[Forms]![VINVerify]![seekvin]</Argu"
                                "ment><Argument Name=\"Expression\">Null</Argument></Action><Action Name=\"SetVal"
                                "ue\"><Argument Name=\"It"
                        End
                        Begin
                            Comment ="_AXL:em\">[Forms]![VINVerify]![MY]</Argument><Argument Name=\"Expression\">Null<"
                                "/Argument></Action><Action Name=\"SetValue\"/></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =6135
                    LayoutCachedTop =15
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =390
                End
            End
        End
        Begin Section
            Height =2160
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1008
                    Top =156
                    Width =3804
                    Height =348
                    ColumnWidth =5415
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1008
                    LayoutCachedTop =156
                    LayoutCachedWidth =4812
                    LayoutCachedHeight =504
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =156
                            Width =810
                            Height =348
                            Name ="UnitVIN_Label"
                            Caption ="VIN"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =120
                            LayoutCachedTop =156
                            LayoutCachedWidth =930
                            LayoutCachedHeight =504
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2856
                    Top =1728
                    Width =732
                    Height =312
                    TabIndex =10
                    Name ="UnitVINLocked"
                    ControlSource ="UnitVINLocked"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2856
                    LayoutCachedTop =1728
                    LayoutCachedWidth =3588
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1008
                            Top =1728
                            Width =1776
                            Height =312
                            Name ="Label13"
                            Caption ="VIN Locked"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =1008
                            LayoutCachedTop =1728
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =2040
                            RowStart =4
                            RowEnd =4
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
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3660
                    Top =1344
                    Width =1152
                    Height =312
                    ColumnWidth =1965
                    FontWeight =700
                    TabIndex =8
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1344
                    LayoutCachedWidth =4812
                    LayoutCachedHeight =1656
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2856
                            Top =1344
                            Width =732
                            Height =312
                            Name ="Label14"
                            Caption ="Status"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =2856
                            LayoutCachedTop =1344
                            LayoutCachedWidth =3588
                            LayoutCachedHeight =1656
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6456
                    Top =960
                    Width =1476
                    Height =315
                    TabIndex =6
                    Name ="UnitVINVer"
                    ControlSource ="UnitVINVer"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6456
                    LayoutCachedTop =960
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =1275
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4884
                            Top =960
                            Width =1500
                            Height =315
                            Name ="Label18"
                            Caption ="Date Verified"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =4884
                            LayoutCachedTop =960
                            LayoutCachedWidth =6384
                            LayoutCachedHeight =1275
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =5
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6456
                    Top =1344
                    Width =1476
                    Height =312
                    TabIndex =9
                    Name ="UnitMSORecd"
                    ControlSource ="UnitMSORecd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6456
                    LayoutCachedTop =1344
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =1656
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4884
                            Top =1344
                            Width =1500
                            Height =312
                            Name ="Label19"
                            Caption ="MSO Rec'd"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =4884
                            LayoutCachedTop =1344
                            LayoutCachedWidth =6384
                            LayoutCachedHeight =1656
                            RowStart =3
                            RowEnd =3
                            ColumnStart =4
                            ColumnEnd =5
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =156
                    Width =852
                    Height =348
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7080
                    LayoutCachedTop =156
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =504
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6456
                            Top =156
                            Width =552
                            Height =348
                            Name ="Label20"
                            Caption ="YEAR"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =6456
                            LayoutCachedTop =156
                            LayoutCachedWidth =7008
                            LayoutCachedHeight =504
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1008
                    Top =576
                    Width =3804
                    Height =312
                    TabIndex =3
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    StatusBarText ="unit make"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1008
                    LayoutCachedTop =576
                    LayoutCachedWidth =4812
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =576
                            Width =810
                            Height =312
                            Name ="Label21"
                            Caption ="Make"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =120
                            LayoutCachedTop =576
                            LayoutCachedWidth =930
                            LayoutCachedHeight =888
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1008
                    Top =960
                    Width =3804
                    Height =315
                    TabIndex =5
                    Name ="UnitModel"
                    ControlSource ="UnitModel"
                    StatusBarText ="unit model"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1008
                    LayoutCachedTop =960
                    LayoutCachedWidth =4812
                    LayoutCachedHeight =1275
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =960
                            Width =810
                            Height =315
                            Name ="Label22"
                            Caption ="Model"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =120
                            LayoutCachedTop =960
                            LayoutCachedWidth =930
                            LayoutCachedHeight =1275
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6456
                    Top =1728
                    Width =1476
                    Height =312
                    TabIndex =11
                    Name ="UnitMSOCopy"
                    ControlSource ="UnitMSOCopy"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6456
                    LayoutCachedTop =1728
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4884
                            Top =1728
                            Width =1500
                            Height =312
                            Name ="Label23"
                            Caption ="MSO Copy "
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =4884
                            LayoutCachedTop =1728
                            LayoutCachedWidth =6384
                            LayoutCachedHeight =2040
                            RowStart =4
                            RowEnd =4
                            ColumnStart =4
                            ColumnEnd =5
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6456
                    Top =576
                    Width =1476
                    Height =312
                    TabIndex =4
                    Name ="UnitGVW"
                    ControlSource ="UnitGVW"
                    Format ="Standard"
                    StatusBarText ="GVW from MSO"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6456
                    LayoutCachedTop =576
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4884
                            Top =576
                            Width =1500
                            Height =312
                            Name ="Label24"
                            Caption ="GVW:"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =4884
                            LayoutCachedTop =576
                            LayoutCachedWidth =6384
                            LayoutCachedHeight =888
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =5
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5808
                    Top =156
                    Width =576
                    Height =348
                    FontWeight =700
                    TabIndex =1
                    BackColor =7883872
                    ForeColor =-2147483607
                    Name ="TitledState"
                    ControlSource ="TitledState"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5808
                    LayoutCachedTop =156
                    LayoutCachedWidth =6384
                    LayoutCachedHeight =504
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4884
                            Top =156
                            Width =852
                            Height =348
                            Name ="Label26"
                            Caption ="State"
                            GroupTable =1
                            RightPadding =38
                            BottomPadding =38
                            LayoutCachedLeft =4884
                            LayoutCachedTop =156
                            LayoutCachedWidth =5736
                            LayoutCachedHeight =504
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1008
                    Top =1344
                    Width =1776
                    Height =312
                    TabIndex =7
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    Format ="General Number"
                    StatusBarText ="unit make"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1008
                    LayoutCachedTop =1344
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =1656
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =120
                            Top =1344
                            Width =810
                            Height =312
                            Name ="Label28"
                            Caption ="AssetID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =120
                            LayoutCachedTop =1344
                            LayoutCachedWidth =930
                            LayoutCachedHeight =1656
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =120
                    Top =1728
                    Width =810
                    Height =312
                    Name ="EmptyCell79"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =120
                    LayoutCachedTop =1728
                    LayoutCachedWidth =930
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3660
                    Top =1728
                    Width =1152
                    Height =312
                    Name ="EmptyCell82"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =3660
                    LayoutCachedTop =1728
                    LayoutCachedWidth =4812
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
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
' See "FindVIN subform.cls"
