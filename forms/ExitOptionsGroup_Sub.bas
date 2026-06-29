Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18736
    DatasheetFontHeight =11
    ItemSuffix =378
    Left =5130
    Top =11235
    Right =23805
    Bottom =18405
    ShortcutMenuBar ="SortClip"
    OrderBy ="[ExitOptionsGroup_Sub].[Selected]"
    RecSrcDt = Begin
        0x3fb70ccfac1ce540
    End
    RecordSource ="SELECT * FROM ExitOptionsGroup_Local;"
    Caption ="Exit Option Group Profile"
    DatasheetFontName ="Calibri"
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
            BorderColor =-2147483609
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BackColor =-2147483613
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
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
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
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
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
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
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
            Height =555
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    SizeMode =1
                    PictureType =2
                    Top =360
                    Width =18045
                    Height =195
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =360
                    LayoutCachedWidth =18045
                    LayoutCachedHeight =555
                    TabIndex =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =10872
                    Top =240
                    Width =1128
                    Height =210
                    FontSize =9
                    Name ="UnitFinalCost_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10872
                    LayoutCachedTop =240
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =450
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    Left =36
                    Top =240
                    Width =780
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    LayoutCachedLeft =36
                    LayoutCachedTop =240
                    LayoutCachedWidth =816
                    LayoutCachedHeight =450
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    Left =876
                    Top =240
                    Width =888
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitUnitNum_Label"
                    Caption ="Unit #"
                    GroupTable =3
                    LayoutCachedLeft =876
                    LayoutCachedTop =240
                    LayoutCachedWidth =1764
                    LayoutCachedHeight =450
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =1836
                    Top =240
                    Width =636
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitStatus_Label"
                    Caption ="Status"
                    GroupTable =3
                    LayoutCachedLeft =1836
                    LayoutCachedTop =240
                    LayoutCachedWidth =2472
                    LayoutCachedHeight =450
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6096
                    Top =240
                    Width =1692
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitVIN_Label"
                    Caption ="Option(s) Selected"
                    GroupTable =3
                    LayoutCachedLeft =6096
                    LayoutCachedTop =240
                    LayoutCachedWidth =7788
                    LayoutCachedHeight =450
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4668
                    Top =240
                    Width =648
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitModel_Label"
                    Caption ="to EOL"
                    GroupTable =3
                    LayoutCachedLeft =4668
                    LayoutCachedTop =240
                    LayoutCachedWidth =5316
                    LayoutCachedHeight =450
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =3
                    Left =7848
                    Top =240
                    Width =2988
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitCOASent_Label"
                    Caption ="New Schedule | Group"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =7848
                    LayoutCachedTop =240
                    LayoutCachedWidth =10836
                    LayoutCachedHeight =450
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =15936
                    Top =240
                    Width =36
                    Height =210
                    Name ="EmptyCell171"
                    GroupTable =3
                    RightPadding =38
                    GridlineColor =0
                    LayoutCachedLeft =15936
                    LayoutCachedTop =240
                    LayoutCachedWidth =15972
                    LayoutCachedHeight =450
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =15828
                    Top =240
                    Width =36
                    Height =210
                    Name ="EmptyCell172"
                    GroupTable =3
                    RightPadding =38
                    GridlineColor =0
                    LayoutCachedLeft =15828
                    LayoutCachedTop =240
                    LayoutCachedWidth =15864
                    LayoutCachedHeight =450
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =5376
                    Top =240
                    Width =648
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label210"
                    Caption ="to Exch"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =5376
                    LayoutCachedTop =240
                    LayoutCachedWidth =6024
                    LayoutCachedHeight =450
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4860
                    Top =60
                    Width =945
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label212"
                    Caption ="┌─ Days ─┐ "
                    RightPadding =38
                    LayoutCachedLeft =4860
                    LayoutCachedTop =60
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2532
                    Top =240
                    Width =2076
                    Height =210
                    FontSize =9
                    ForeColor =6108695
                    Name ="Label517"
                    Caption ="┌─ Add/Edit Unit─┐"
                    GroupTable =3
                    LayoutCachedLeft =2532
                    LayoutCachedTop =240
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =450
                    ColumnStart =3
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =8160
                    Width =7035
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label336"
                    Caption ="┌──────────────── ExchangeIT Equipment ──────────────────────────────────────┐"
                    RightPadding =38
                    LayoutCachedLeft =8160
                    LayoutCachedWidth =15195
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Width =1260
                    Height =270
                    FontSize =9
                    Name ="GroupID"
                    ControlSource ="GroupID"
                    StatusBarText ="schedule number starts with 1"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedWidth =1260
                    LayoutCachedHeight =270
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13932
                    Top =240
                    Width =876
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label370"
                    Caption ="Accepeted"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =13932
                    LayoutCachedTop =240
                    LayoutCachedWidth =14808
                    LayoutCachedHeight =450
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12036
                    Top =240
                    Width =876
                    Height =210
                    FontSize =9
                    Name ="Label665"
                    Caption ="Est Delv"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =12036
                    LayoutCachedTop =240
                    LayoutCachedWidth =12912
                    LayoutCachedHeight =450
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14880
                    Top =240
                    Width =876
                    Height =210
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label735"
                    Caption ="In Service"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =14880
                    LayoutCachedTop =240
                    LayoutCachedWidth =15756
                    LayoutCachedHeight =450
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12984
                    Top =240
                    Width =876
                    Height =210
                    FontSize =9
                    Name ="Label737"
                    Caption ="Act Delv"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =12984
                    LayoutCachedTop =240
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =450
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
            End
        End
        Begin Section
            Height =240
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4668
                    Width =648
                    ColumnWidth =1410
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =7
                    Name ="UnitModel"
                    ControlSource ="=IIf([DaysToEOL]<0,0,[DaysToEOL])"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4668
                    LayoutCachedWidth =5316
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =36
                    Width =780
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =36
                    LayoutCachedWidth =816
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15828
                    Width =36
                    ColumnWidth =3000
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =16
                    Name ="SGrpID"
                    ControlSource ="GroupId"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =15828
                    LayoutCachedWidth =15864
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1836
                    Width =636
                    ColumnWidth =840
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =2
                    BorderColor =14211288
                    Name ="UnitStatus"
                    ControlSource ="uStatus"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1836
                    LayoutCachedWidth =2472
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6096
                    Width =1692
                    ColumnWidth =2100
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =9
                    Name ="UnitVIN"
                    ControlSource ="Selected"
                    StatusBarText ="VIN number must be unique in table"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6096
                    LayoutCachedWidth =7788
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =876
                    Width =888
                    ColumnWidth =1500
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =876
                    LayoutCachedWidth =1764
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7848
                    Width =2988
                    FontSize =9
                    TabIndex =10
                    BorderColor =14211288
                    ForeColor =8210719
                    Name ="NewSchGrp"
                    ControlSource ="=IIf(IsNull([newSchedule]),Null,[NewClient] & \" | \" & [NewSchedule] & \" | \" "
                        "& [NewGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =7848
                    LayoutCachedWidth =10836
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10872
                    Width =1128
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =11
                    Name ="NewUnit"
                    ControlSource ="NewAssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10872
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15936
                    Width =36
                    FontSize =9
                    TabIndex =17
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =15936
                    LayoutCachedWidth =15972
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2532
                    Width =468
                    Height =240
                    FontSize =8
                    TabIndex =3
                    Name ="ModifyExch"
                    Caption ="EXCH"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units ExchangeIT Options"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2532
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8872517
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5376
                    Width =648
                    FontSize =9
                    TabIndex =8
                    Name ="Text204"
                    ControlSource ="=IIf([DaysToExch]<0,0,[DaysToExch])"
                    StatusBarText ="unit model"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =5376
                    LayoutCachedWidth =6024
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3588
                    Width =492
                    Height =240
                    FontSize =8
                    TabIndex =5
                    Name ="ModifyTERM"
                    Caption ="TERM"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units Termation Options"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3588
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =790406
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =790406
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3060
                    Width =468
                    Height =240
                    FontSize =8
                    TabIndex =4
                    Name ="ModifyEXTD"
                    Caption ="EXTD"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units Lease Extension Options"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3060
                    LayoutCachedWidth =3528
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =9211020
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4140
                    Width =468
                    Height =240
                    FontSize =8
                    TabIndex =6
                    Name ="ModifyHIST"
                    Caption ="HIST"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="See this Units EXIT History"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =50.0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =12566463
                    PressedColor =9211020
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13932
                    Width =876
                    FontSize =9
                    TabIndex =14
                    Name ="LastModified"
                    ControlSource ="nAccepted"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =13932
                    LayoutCachedWidth =14808
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12036
                    Width =876
                    FontSize =9
                    TabIndex =12
                    Name ="Text659"
                    ControlSource ="=[nEstDelv]"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =12036
                    LayoutCachedWidth =12912
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12984
                    Width =876
                    FontSize =9
                    TabIndex =13
                    Name ="Text743"
                    ControlSource ="[nActualDelv]"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =12984
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Width =876
                    FontSize =9
                    TabIndex =15
                    Name ="Text733"
                    ControlSource ="nInServ"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =14880
                    LayoutCachedWidth =15756
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GroupTable =3
                End
            End
        End
        Begin FormFooter
            Height =15
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =15
                    Height =0
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =15
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =15
                End
            End
        End
    End
End
CodeBehindForm
' See "ExitOptionsGroup_Sub.cls"
