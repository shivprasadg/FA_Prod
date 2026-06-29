Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15840
    DatasheetFontHeight =11
    ItemSuffix =218
    Left =1395
    Top =8760
    Right =19530
    Bottom =11625
    RecSrcDt = Begin
        0x155e20940d02e540
    End
    RecordSource ="SELECT * From vw_TaxStatesExceptionsQ WHERE TaxTypeID =2 AND TaxStatesID=11 ORDE"
        "R BY StartsOn DESC;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =480
            BackColor =15527148
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =90
                    Width =4065
                    Height =315
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="StateNameLabel"
                    ControlSource ="=\"Exception to: \" & DLookUp(\"TaxFullName\",\"TaxType\",\"TaxTypeID=\" & [TaxT"
                        "ypeId] & \"\")"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedTop =90
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =405
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =9090
                    Top =60
                    Width =3165
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label242"
                    Caption ="┌──Unit ───────────────────────┐"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =9090
                    LayoutCachedTop =60
                    LayoutCachedWidth =12255
                    LayoutCachedHeight =300
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4934
                    Width =1080
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Starts \015\012ON"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4934
                    LayoutCachedWidth =6014
                    LayoutCachedHeight =480
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3180
                    Width =120
                    Height =210
                    ColumnOrder =1
                    FontSize =8
                    TabIndex =1
                    BorderColor =12632256
                    ForeColor =0
                    Name ="txtCurrent"
                    GridlineColor =0

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =210
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4079
                    Width =835
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="TaxNumberLabel"
                    Caption ="Excp Tax"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4079
                    LayoutCachedWidth =4914
                    LayoutCachedHeight =240
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4079
                    Top =240
                    Width =835
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label206"
                    Caption ="Rate %"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4079
                    LayoutCachedTop =240
                    LayoutCachedWidth =4914
                    LayoutCachedHeight =480
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =6014
                    Width =1080
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label210"
                    Caption ="Ends\015\012ON"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6014
                    LayoutCachedWidth =7094
                    LayoutCachedHeight =480
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =15210
                    Top =240
                    Width =614
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label218"
                    Caption ="Active"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =15210
                    LayoutCachedTop =240
                    LayoutCachedWidth =15824
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =8160
                    Width =765
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label220"
                    Caption =" Replaces Base"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8160
                    LayoutCachedWidth =8925
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =8926
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label223"
                    Caption ="Trac"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8926
                    LayoutCachedTop =255
                    LayoutCachedWidth =9459
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =9992
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label226"
                    Caption ="Trlr"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =9992
                    LayoutCachedTop =255
                    LayoutCachedWidth =10525
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =9459
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label229"
                    Caption ="Str"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =9459
                    LayoutCachedTop =255
                    LayoutCachedWidth =9992
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =12975
                    Top =255
                    Width =555
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label232"
                    Caption ="NAICS"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =12975
                    LayoutCachedTop =255
                    LayoutCachedWidth =13530
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =12325
                    Top =255
                    Width =585
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label235"
                    Caption ="Term↑ "
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =12325
                    LayoutCachedTop =255
                    LayoutCachedWidth =12910
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =10525
                    Top =255
                    Width =900
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label236"
                    Caption ="Weight ↑"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10525
                    LayoutCachedTop =255
                    LayoutCachedWidth =11425
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7094
                    Width =1080
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label240"
                    Caption ="Max\015\012  Amount $"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7094
                    LayoutCachedWidth =8174
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =11425
                    Top =255
                    Width =900
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label244"
                    Caption ="Weight ↓"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =11425
                    LayoutCachedTop =255
                    LayoutCachedWidth =12325
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =12325
                    Top =60
                    Width =585
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label246"
                    Caption ="Lease"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =12325
                    LayoutCachedTop =60
                    LayoutCachedWidth =12910
                    LayoutCachedHeight =300
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =13575
                    Top =255
                    Width =1005
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label249"
                    Caption ="Forms-Docs"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =13575
                    LayoutCachedTop =255
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =480
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =298
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =8191
                    Width =750
                    Height =255
                    BackColor =12566463
                    Name ="Box238"
                    LayoutCachedLeft =8191
                    LayoutCachedWidth =8941
                    LayoutCachedHeight =255
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =45
                    Width =4020
                    Height =255
                    FontSize =10
                    ForeColor =0
                    Name ="ExceptionDescription"
                    ControlSource ="[ExceptionDescription]"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Exception Description"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =45
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4934
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="StartsOn"
                    ControlSource ="StartsOn"
                    Format ="Short Date"
                    ControlTipText ="The Date This Exception Starts"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4934
                    LayoutCachedWidth =6014
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3900
                    Width =153
                    Height =231
                    FontSize =8
                    TabIndex =2
                    Name ="Id"
                    ControlSource ="Id"
                    TopPadding =0

                    LayoutCachedLeft =3900
                    LayoutCachedWidth =4053
                    LayoutCachedHeight =231
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =15419
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =3
                    Name ="IsActive"
                    ControlSource ="IsActive"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15419
                    LayoutCachedTop =29
                    LayoutCachedWidth =15629
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Width =138
                    Height =231
                    FontSize =8
                    TabIndex =4
                    Name ="TaxStatesID"
                    ControlSource ="TaxStatesID"
                    TopPadding =0

                    LayoutCachedLeft =3495
                    LayoutCachedWidth =3633
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6014
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="EndsOn"
                    ControlSource ="EndsOn"
                    Format ="Short Date"
                    ControlTipText ="The Date This Exception Expires"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6014
                    LayoutCachedWidth =7094
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =255
                    Width =15840
                    Height =43
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490044005d003d005b00740078007400430075007200720065006e007400 ,
                        0x5d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =255
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =298
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00110000005b00 ,
                        0x490044005d003d005b00740078007400430075007200720065006e0074005d00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4079
                    Width =835
                    Height =255
                    FontSize =10
                    TabIndex =7
                    BackColor =11525325
                    BorderColor =9211020
                    ForeColor =0
                    Name ="TaxNumber"
                    ControlSource ="TaxRate"
                    Format ="Percent"
                    ValidationRule ="<=1 And >=-1"
                    ValidationText ="Tax Rates cannot be > 100%"
                    ControlTipText ="The percentage of the Tax Rate Exception"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4079
                    LayoutCachedWidth =4914
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3195
                    Width =138
                    Height =231
                    FontSize =8
                    TabIndex =8
                    Name ="TaxTypeId"
                    ControlSource ="TaxTypeId"
                    TopPadding =0

                    LayoutCachedLeft =3195
                    LayoutCachedWidth =3333
                    LayoutCachedHeight =231
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =215
                    Left =8491
                    Top =43
                    Width =210
                    Height =195
                    TabIndex =9
                    Name ="Check219"
                    ControlSource ="TaxRule"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8491
                    LayoutCachedTop =43
                    LayoutCachedWidth =8701
                    LayoutCachedHeight =238
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =9091
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =10
                    Name ="AppliesToTractors"
                    ControlSource ="AppliesToTractors"
                    ControlTipText ="This Exception Applies To Tractors"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9091
                    LayoutCachedTop =29
                    LayoutCachedWidth =9301
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =10171
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =11
                    Name ="AppliesToTrailers"
                    ControlSource ="AppliesToTrailers"
                    ControlTipText ="This Exception Applies To Trailers"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10171
                    LayoutCachedTop =29
                    LayoutCachedWidth =10381
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =9631
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =12
                    Name ="AppliesToStraightTrucks"
                    ControlSource ="AppliesToStraightTrucks"
                    ControlTipText ="This Exception Applies To Straight Trucks"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9631
                    LayoutCachedTop =29
                    LayoutCachedWidth =9841
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10525
                    Width =900
                    Height =255
                    FontSize =10
                    TabIndex =13
                    BorderColor =9211020
                    ForeColor =0
                    Name ="WeightLimitMinLbs"
                    ControlSource ="WeightLimitMinLbs"
                    Format ="Standard"
                    ControlTipText ="If the Weight is GREATER than this Value"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10525
                    LayoutCachedWidth =11425
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12325
                    Width =585
                    Height =255
                    FontSize =10
                    TabIndex =14
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Text234"
                    ControlSource ="TermExceedsMonths"
                    ControlTipText ="If the Lease Term is Longer than this value"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12325
                    LayoutCachedWidth =12910
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7094
                    Width =1080
                    Height =255
                    FontSize =10
                    TabIndex =15
                    BorderColor =9211020
                    ForeColor =0
                    Name ="TaxMaxValue"
                    ControlSource ="TaxMaxValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Tax cannot excced this Amount"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7094
                    LayoutCachedWidth =8174
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2445
                    Width =138
                    Height =231
                    FontSize =8
                    TabIndex =16
                    Name ="TaxCitiesID"
                    ControlSource ="TaxCitiesID"
                    TopPadding =0

                    LayoutCachedLeft =2445
                    LayoutCachedWidth =2583
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2730
                    Width =138
                    Height =231
                    FontSize =8
                    TabIndex =17
                    Name ="TaxCountiesID"
                    ControlSource ="TaxCountiesID"
                    TopPadding =0

                    LayoutCachedLeft =2730
                    LayoutCachedWidth =2868
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11425
                    Width =900
                    Height =255
                    FontSize =10
                    TabIndex =18
                    BorderColor =9211020
                    ForeColor =0
                    Name ="WeightLimitMaxLbs"
                    ControlSource ="WeightLimitMaxLbs"
                    Format ="Standard"
                    ControlTipText ="If the Weight is LESS than this Value"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11425
                    LayoutCachedWidth =12325
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13590
                    Width =495
                    Height =255
                    FontSize =10
                    TabIndex =19
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Web1"
                    ControlSource ="Web1"
                    ControlTipText ="If valid, this is a Web Link explaining this Exception"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13590
                    LayoutCachedWidth =14085
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14085
                    Width =495
                    Height =255
                    FontSize =10
                    TabIndex =20
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Web2"
                    ControlSource ="Web2"
                    ControlTipText ="If valid, this is a Web Link for an Application to this Exception"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14085
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =14700
                    Width =465
                    Height =256
                    FontSize =8
                    TabIndex =21
                    Name ="EditMe"
                    Caption ="Edit"
                    ControlTipText ="Edit This Exception"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =14700
                    LayoutCachedWidth =15165
                    LayoutCachedHeight =256
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8965045
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =9211020
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12975
                    Width =585
                    Height =255
                    FontSize =10
                    TabIndex =22
                    BorderColor =9211020
                    ForeColor =0
                    Name ="NAICS"
                    ControlSource ="=IIf(Len([IncludedNAICS])>2,\"Y\",Null)"
                    ControlTipText ="Are there NAICS Codes assigned to the Exception? YN"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12975
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9405
                    Width =138
                    Height =231
                    FontSize =8
                    TabIndex =23
                    Name ="TaxScript"
                    ControlSource ="TaxScript"
                    TopPadding =0

                    LayoutCachedLeft =9405
                    LayoutCachedWidth =9543
                    LayoutCachedHeight =231
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =14211288
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "Tax_Add_ExceptionSub.cls"
