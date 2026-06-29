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
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =242
    Left =2475
    Top =8490
    Right =18060
    Bottom =12120
    RecSrcDt = Begin
        0x49c446442cfbe440
    End
    RecordSource ="SELECT EXP.ID, EXP.ExceptionDescription, EXP.TaxRule, EXP.TaxRate, EXP.FormURL1,"
        " EXP.FormURL2, EXP.IncludedNAICS, EXP.AppliesToTractors, EXP.AppliesToTrailers, "
        "EXP.AppliesToStraightTrucks, EXP.WeightLimitMinLbs, EXP.WeightLimitMaxLbs, EXP.T"
        "ermExceedsMonths, EXP.ForHireOnly, EXP.StartsOn, EXP.EndsOn, EXP.TaxCountyID, EX"
        "P.TaxTypeId, TaxType.TaxFullName, IIf(Not IsNull([exp].[StartsOn]) And IsNull([e"
        "xp].[EndsOn]),-1,0) AS IsActive, TaxType.TaxAbbreviation, TaxType.TaxDescription"
        ", EXP.TaxMaxValue FROM TaxCountiesExceptions AS EXP LEFT JOIN TaxType ON EXP.Tax"
        "TypeID = TaxType.Id;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3180
                    Width =120
                    Height =210
                    ColumnOrder =0
                    FontSize =8
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =90
                    Width =4065
                    Height =315
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="StateNameLabel"
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =12345
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
                    LayoutCachedLeft =12345
                    LayoutCachedTop =240
                    LayoutCachedWidth =12959
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =7094
                    Width =765
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label220"
                    Caption =" Replaces Base"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7094
                    LayoutCachedWidth =7859
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =7860
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label223"
                    Caption ="Trac"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7860
                    LayoutCachedTop =255
                    LayoutCachedWidth =8393
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =8926
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label226"
                    Caption ="Trlr"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8926
                    LayoutCachedTop =255
                    LayoutCachedWidth =9459
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =8393
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label229"
                    Caption ="Str"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8393
                    LayoutCachedTop =255
                    LayoutCachedWidth =8926
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =9459
                    Top =255
                    Width =533
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label232"
                    Caption ="4 Hire"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =9459
                    LayoutCachedTop =255
                    LayoutCachedWidth =9992
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =9992
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
                    LayoutCachedLeft =9992
                    LayoutCachedTop =255
                    LayoutCachedWidth =10577
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =10577
                    Top =255
                    Width =900
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label236"
                    Caption ="Weight ↑"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10577
                    LayoutCachedTop =255
                    LayoutCachedWidth =11477
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =11477
                    Top =255
                    Width =835
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label240"
                    Caption ="Max $"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =11477
                    LayoutCachedTop =255
                    LayoutCachedWidth =12312
                    LayoutCachedHeight =480
                    ColumnStart =7
                    ColumnEnd =9
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8040
                    Top =60
                    Width =3285
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label242"
                    Caption ="┌─────────── Applies to ───────────┐"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8040
                    LayoutCachedTop =60
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =300
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
            AlternateBackColor =16777215
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =7125
                    Width =750
                    Height =255
                    BackColor =12566463
                    Name ="Box238"
                    LayoutCachedLeft =7125
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =255
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =45
                    Width =4020
                    Height =255
                    FontSize =10
                    ForeColor =0
                    Name ="Text214"
                    ControlSource ="=[ExceptionDescription] & \" (\" & [TaxAbbreviation] & \")\""
                    ControlTipText ="Double-Click to Edit Event"
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
                    ControlTipText ="Double-Click to Edit Event"
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
                    Left =12586
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

                    LayoutCachedLeft =12586
                    LayoutCachedTop =29
                    LayoutCachedWidth =12796
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
                    ControlTipText ="Double-Click to Edit Event"
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
                    Width =12960
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
                    LayoutCachedWidth =12960
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
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4079
                    Width =835
                    Height =255
                    FontSize =10
                    TabIndex =7
                    BorderColor =9211020
                    ForeColor =0
                    Name ="TaxNumber"
                    ControlSource ="TaxRate"
                    Format ="Percent"
                    ValidationRule ="<=1 And >=-1"
                    ValidationText ="Tax Rates cannot be > 100%"
                    ControlTipText ="The percentage of the Tax Rate"
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
                    OverlapFlags =215
                    Left =7425
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

                    LayoutCachedLeft =7425
                    LayoutCachedTop =43
                    LayoutCachedWidth =7635
                    LayoutCachedHeight =238
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8025
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =10
                    Name ="AppliesToTractors"
                    ControlSource ="AppliesToTractors"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8025
                    LayoutCachedTop =29
                    LayoutCachedWidth =8235
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9105
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =11
                    Name ="AppliesToTrailers"
                    ControlSource ="AppliesToTrailers"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9105
                    LayoutCachedTop =29
                    LayoutCachedWidth =9315
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8565
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =12
                    Name ="AppliesToStraightTrucks"
                    ControlSource ="AppliesToStraightTrucks"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8565
                    LayoutCachedTop =29
                    LayoutCachedWidth =8775
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9645
                    Top =29
                    Width =210
                    Height =225
                    TabIndex =13
                    Name ="Check233"
                    ControlSource ="ForHireOnly"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9645
                    LayoutCachedTop =29
                    LayoutCachedWidth =9855
                    LayoutCachedHeight =254
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10577
                    Width =990
                    Height =255
                    FontSize =10
                    TabIndex =14
                    BorderColor =9211020
                    ForeColor =0
                    Name ="IfWeightLimitExceedsLbs"
                    ControlSource ="WeightLimitMinLbs"
                    Format ="Standard"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10577
                    LayoutCachedWidth =11567
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
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9992
                    Width =585
                    Height =255
                    FontSize =10
                    TabIndex =15
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Text234"
                    ControlSource ="TermExceedsMonths"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9992
                    LayoutCachedWidth =10577
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
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11567
                    Width =835
                    Height =255
                    FontSize =10
                    TabIndex =16
                    BorderColor =9211020
                    ForeColor =0
                    Name ="TaxMaxValue"
                    ControlSource ="TaxMaxValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11567
                    LayoutCachedWidth =12402
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
' See "Tax_CountyExceptionsSub.cls"
