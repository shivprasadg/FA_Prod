Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21893
    DatasheetFontHeight =11
    ItemSuffix =853
    Left =5700
    Top =6105
    Right =26475
    Bottom =10020
    ShortcutMenuBar ="SortClip"
    OrderBy ="[ExitOptionsQryExtend].[UnitID]"
    RecSrcDt = Begin
        0xcd91465b8cd0e540
    End
    RecordSource ="ExitOptionsQryExtend"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1035
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =645
                    Width =20880
                    Height =345
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =645
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =990
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =345
                    Width =20880
                    Height =285
                    BorderColor =1643706
                    Name ="Box47"
                    LayoutCachedTop =345
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =630
                    BackThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =36
                    Top =696
                    Width =828
                    Height =276
                    FontSize =10
                    ForeColor =8210719
                    Name ="Label20"
                    Caption ="AssetID"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =696
                    LayoutCachedWidth =864
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =12
                    Width =756
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="TermSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5880
                    LayoutCachedTop =12
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =252
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =36
                    Top =12
                    Width =828
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="send2Excel"
                    Caption ="Excel"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120

                    LayoutCachedLeft =36
                    LayoutCachedTop =12
                    LayoutCachedWidth =864
                    LayoutCachedHeight =252
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =36
                    Top =372
                    Width =828
                    Height =288
                    Name ="EmptyCell69"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =372
                    LayoutCachedWidth =864
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5880
                    Top =372
                    Width =756
                    Height =288
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Months"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =372
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =12
                    Width =1068
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="PurposedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4740
                    LayoutCachedTop =12
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =252
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6708
                    Top =12
                    Width =1068
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="ActualDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6708
                    LayoutCachedTop =12
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =252
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8952
                    Top =696
                    Width =3900
                    Height =276
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =15
                    Name ="NotesCbo"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Notes All listed units"
                    GroupTable =1

                    LayoutCachedLeft =8952
                    LayoutCachedTop =696
                    LayoutCachedWidth =12852
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =8952
                    Top =372
                    Width =3900
                    Height =288
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label504"
                    Caption =" ExtendIT Note"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8952
                    LayoutCachedTop =372
                    LayoutCachedWidth =12852
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8952
                    Top =12
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="NotesSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8952
                    LayoutCachedTop =12
                    LayoutCachedWidth =10104
                    LayoutCachedHeight =252
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =10164
                    Top =12
                    Width =2688
                    Name ="EmptyCell508"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =10164
                    LayoutCachedTop =12
                    LayoutCachedWidth =12852
                    LayoutCachedHeight =252
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4740
                    Top =696
                    Width =1068
                    Height =276
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =11
                    Name ="StartOrPurposedDateCBO"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Lease Start Date for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =4740
                    LayoutCachedTop =696
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6708
                    Top =696
                    Width =1068
                    Height =276
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =13
                    Name ="EndOrActualDateCBO"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Lease End Date for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =6708
                    LayoutCachedTop =696
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =5880
                    Top =696
                    Width =756
                    Height =276
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =12
                    Name ="TermCBO"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30"
                        ";31;32;33;34;35;36"
                    ColumnWidths ="864"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Lease Term for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =5880
                    LayoutCachedTop =696
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4740
                    Top =372
                    Width =1068
                    Height =288
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label271"
                    Caption =" Start Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4740
                    LayoutCachedTop =372
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6708
                    Top =372
                    Width =1068
                    Height =288
                    FontSize =10
                    ForeColor =16777215
                    Name ="Text520"
                    Caption ="End Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6708
                    LayoutCachedTop =372
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =924
                    Top =12
                    Width =876
                    Name ="EmptyCell561"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =924
                    LayoutCachedTop =12
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =924
                    Top =372
                    Width =876
                    Height =288
                    Name ="EmptyCell571"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =924
                    LayoutCachedTop =372
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1872
                    Top =372
                    Width =1152
                    Height =288
                    FontSize =10
                    Name ="Label583"
                    Caption =" Extd Type"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1872
                    LayoutCachedTop =372
                    LayoutCachedWidth =3024
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7848
                    Top =372
                    Width =1032
                    Height =288
                    FontSize =10
                    Name ="Label584"
                    Caption ="Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7848
                    LayoutCachedTop =372
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7848
                    Top =696
                    Width =1032
                    Height =276
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =14
                    Name ="RentCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Lease Rent for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =7848
                    LayoutCachedTop =696
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7848
                    Top =12
                    Width =1032
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="RentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7848
                    LayoutCachedTop =12
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =252
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1872
                    Top =12
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="OptionTypeDescSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1872
                    LayoutCachedTop =12
                    LayoutCachedWidth =3024
                    LayoutCachedHeight =252
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =20184
                    Top =372
                    Width =876
                    Height =288
                    FontSize =10
                    Name ="Label623"
                    Caption ="Residual"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20184
                    LayoutCachedTop =372
                    LayoutCachedWidth =21060
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =21132
                    Top =696
                    Width =36
                    Height =276
                    FontSize =10
                    Name ="Label631"
                    Caption ="Orig RV$"
                    GroupTable =1
                    LayoutCachedLeft =21132
                    LayoutCachedTop =696
                    LayoutCachedWidth =21168
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20184
                    Top =696
                    Width =876
                    Height =276
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =18
                    Name ="ResidualCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="Set a New Residual for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =20184
                    LayoutCachedTop =696
                    LayoutCachedWidth =21060
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20184
                    Top =12
                    Width =876
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    ForeColor =8210719
                    Name ="ResidualSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20184
                    LayoutCachedTop =12
                    LayoutCachedWidth =21060
                    LayoutCachedHeight =252
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverThemeColorIndex =1
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeColor =0
                    PressedForeColor =8210719
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3096
                    Top =12
                    Width =1572
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="OptionSubTypeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the value below for all Units shown on form"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3096
                    LayoutCachedTop =12
                    LayoutCachedWidth =4668
                    LayoutCachedHeight =252
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =3096
                    Top =696
                    Width =1572
                    Height =276
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =10
                    Name ="OptionSubTypeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OptionSubType, OptionSubTypeDesc, ExitOptionType FROM LookUp_ExitOptionTy"
                        "pes WHERE ExitOptionType='Extend';"
                    ColumnWidths ="1728;2160"
                    GroupTable =1

                    LayoutCachedLeft =3096
                    LayoutCachedTop =696
                    LayoutCachedWidth =4668
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3096
                    Top =372
                    Width =1572
                    Height =288
                    FontSize =10
                    Name ="LOlbl"
                    Caption ="ExtendIT Sub Type"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3096
                    LayoutCachedTop =372
                    LayoutCachedWidth =4668
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14616
                    Top =372
                    Width =1116
                    Height =288
                    Name ="EmptyCell716"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14616
                    LayoutCachedTop =372
                    LayoutCachedWidth =15732
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12924
                    Top =12
                    Width =468
                    Name ="EmptyCell719"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =12924
                    LayoutCachedTop =12
                    LayoutCachedWidth =13392
                    LayoutCachedHeight =252
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14616
                    Top =12
                    Width =1116
                    Name ="EmptyCell728"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =14616
                    LayoutCachedTop =12
                    LayoutCachedWidth =15732
                    LayoutCachedHeight =252
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14004
                    Top =12
                    Width =480
                    Name ="EmptyCell742"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =14004
                    LayoutCachedTop =12
                    LayoutCachedWidth =14484
                    LayoutCachedHeight =252
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =13464
                    Top =12
                    Width =468
                    Name ="EmptyCell747"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =13464
                    LayoutCachedTop =12
                    LayoutCachedWidth =13932
                    LayoutCachedHeight =252
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =15804
                    Top =696
                    Width =1116
                    Height =276
                    FontSize =8
                    TabIndex =16
                    Name ="SaveUpdateMe"
                    Caption ="Save/Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save Update Changes"
                    GroupTable =1

                    LayoutCachedLeft =15804
                    LayoutCachedTop =696
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderColor =16777215
                    HoverColor =12428174
                    PressedColor =8210719
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =12924
                            Top =696
                            Width =2808
                            Height =276
                            FontSize =10
                            Name ="Label764"
                            Caption ="  ↓ Additional Extension"
                            GroupTable =1
                            LayoutCachedLeft =12924
                            LayoutCachedTop =696
                            LayoutCachedWidth =15732
                            LayoutCachedHeight =972
                            RowStart =2
                            RowEnd =2
                            ColumnStart =10
                            ColumnEnd =14
                            LayoutGroup =1
                            BackThemeColorIndex =4
                            ForeThemeColorIndex =2
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =12924
                    Top =372
                    Width =1560
                    Height =288
                    Name ="EmptyCell774"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12924
                    LayoutCachedTop =372
                    LayoutCachedWidth =14484
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =21132
                    Top =12
                    Width =36
                    Name ="EmptyCell805"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =21132
                    LayoutCachedTop =12
                    LayoutCachedWidth =21168
                    LayoutCachedHeight =252
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =21132
                    Top =372
                    Width =36
                    Height =288
                    Name ="EmptyCell806"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21132
                    LayoutCachedTop =372
                    LayoutCachedWidth =21168
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17340
                    Top =12
                    Width =2772
                    Name ="EmptyCell821"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =17340
                    LayoutCachedTop =12
                    LayoutCachedWidth =20112
                    LayoutCachedHeight =252
                    ColumnStart =17
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17340
                    Top =372
                    Width =2772
                    Height =288
                    Name ="EmptyCell822"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17340
                    LayoutCachedTop =372
                    LayoutCachedWidth =20112
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16980
                    Top =12
                    Width =300
                    Name ="EmptyCell825"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =16980
                    LayoutCachedTop =12
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =252
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16980
                    Top =372
                    Width =300
                    Height =288
                    Name ="EmptyCell826"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16980
                    LayoutCachedTop =372
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15804
                    Top =12
                    Width =1116
                    Name ="EmptyCell829"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =15804
                    LayoutCachedTop =12
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =252
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15804
                    Top =372
                    Width =1116
                    Height =288
                    Name ="EmptyCell830"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15804
                    LayoutCachedTop =372
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =18780
                    Top =696
                    Width =660
                    Height =276
                    FontSize =10
                    TabIndex =17
                    ForeColor =16777215
                    Name ="DeleteAll"
                    Caption ="Õ"
                    OnClick ="[Event Procedure]"
                    FontName ="Wingdings"
                    ControlTipText ="Delete All Options"
                    GroupTable =1

                    LayoutCachedLeft =18780
                    LayoutCachedTop =696
                    LayoutCachedWidth =19440
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =1643706
                    PressedColor =1643706
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =16980
                    Top =696
                    Width =300
                    Height =276
                    Name ="EmptyCell835"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =16980
                    LayoutCachedTop =696
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17340
                    Top =696
                    Width =1380
                    Height =276
                    Name ="EmptyCell838"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =17340
                    LayoutCachedTop =696
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19500
                    Top =696
                    Width =612
                    Height =276
                    Name ="EmptyCell839"
                    GroupTable =1
                    LayoutCachedLeft =19500
                    LayoutCachedTop =696
                    LayoutCachedWidth =20112
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14544
                    Top =12
                    Width =12
                    Name ="EmptyCell842"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =14544
                    LayoutCachedTop =12
                    LayoutCachedWidth =14556
                    LayoutCachedHeight =252
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14544
                    Top =372
                    Width =12
                    Height =288
                    Name ="EmptyCell843"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14544
                    LayoutCachedTop =372
                    LayoutCachedWidth =14556
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1872
                    Top =696
                    Width =1152
                    Height =276
                    FontSize =10
                    TabIndex =8
                    Name ="OptionTypeDescCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ExitOptionTypes.ExitOptionType FROM LookUp_ExitOptionTypes WHERE ("
                        "((LookUp_ExitOptionTypes.OptionTypeID)>2 Or (LookUp_ExitOptionTypes.OptionTypeID"
                        ")<0)) GROUP BY LookUp_ExitOptionTypes.ExitOptionType;"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    ControlTipText ="Set Extended Lease Service Fee for All listed units"
                    GroupTable =1

                    LayoutCachedLeft =1872
                    LayoutCachedTop =696
                    LayoutCachedWidth =3024
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =924
                            Top =696
                            Width =876
                            Height =276
                            FontSize =10
                            Name ="Label577"
                            Caption ="Unit#"
                            GroupTable =1
                            LayoutCachedLeft =924
                            LayoutCachedTop =696
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =972
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =1
                            BackThemeColorIndex =4
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Top =30
                    Width =0
                    Height =270
                    FontSize =10
                    TabIndex =19
                    Name ="OptionTypeID"
                    ControlSource ="OptionTypeID"

                    LayoutCachedTop =30
                    LayoutCachedHeight =300
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6708
                    Top =36
                    Width =1068
                    Height =276
                    FontSize =10
                    TabIndex =6
                    Name ="EndOrActualDate"
                    ControlSource ="EndOrActualDate"
                    ValidationText ="BO Date is required"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Extended Lease End Date"
                    GroupTable =1

                    LayoutCachedLeft =6708
                    LayoutCachedTop =36
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =828
                    Height =276
                    FontSize =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =864
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =20880
                    Height =59
                    TabIndex =20
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =315
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4740
                    Top =36
                    Width =1068
                    Height =276
                    FontSize =10
                    TabIndex =4
                    Name ="StartOrPurposedDate"
                    ControlSource ="StartOrPurposedDate"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText =" Extended Lease Start Date"
                    GroupTable =1

                    LayoutCachedLeft =4740
                    LayoutCachedTop =36
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8952
                    Top =36
                    Width =3900
                    Height =276
                    FontSize =10
                    TabIndex =8
                    Name ="Notes"
                    ControlSource ="Notes"
                    OnChange ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =8952
                    LayoutCachedTop =36
                    LayoutCachedWidth =12852
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14616
                    Top =36
                    Width =1116
                    Height =276
                    FontSize =8
                    TabIndex =13
                    Name ="UserName"
                    ControlSource ="UserName"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =14616
                    LayoutCachedTop =36
                    LayoutCachedWidth =15732
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15804
                    Top =36
                    Width =1116
                    Height =276
                    FontSize =8
                    TabIndex =14
                    Name ="DatePosted"
                    ControlSource ="DatePosted"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =15804
                    LayoutCachedTop =36
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =924
                    Top =36
                    Width =876
                    Height =276
                    FontSize =10
                    TabIndex =2
                    BackColor =16317438
                    Name ="Text575"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1

                    LayoutCachedLeft =924
                    LayoutCachedTop =36
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7848
                    Top =36
                    Width =1032
                    Height =276
                    FontSize =10
                    TabIndex =7
                    Name ="Rent"
                    ControlSource ="Rent"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Extended Lease Rent"
                    GroupTable =1

                    LayoutCachedLeft =7848
                    LayoutCachedTop =36
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =5880
                    Top =36
                    Width =756
                    Height =276
                    FontSize =10
                    TabIndex =5
                    Name ="Term"
                    ControlSource ="TermMonths"
                    RowSourceType ="Value List"
                    RowSource ="\"1\";\"2\";\"3\";\"4\";\"5\";\"6\";\"7\";\"8\";\"9\";\"10\";\"11\";\"12\";\"13\""
                        ";\"14\";\"15\";\"16\";\"17\";\"18\";\"19\";\"20\";\"21\";\"22\";\"23\";\"24\";\""
                        "25\";\"26\";\"27\";\"28\";\"29\";\"30\";\"31\";\"32\";\"33\";\"34\";\"35\";\"36\""
                        ";\"40\""
                    ColumnWidths ="864"
                    DefaultValue ="=Null"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Tag ="combos"
                    ControlTipText ="Extended Lease Term"
                    GroupTable =1

                    LayoutCachedLeft =5880
                    LayoutCachedTop =36
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20184
                    Top =36
                    Width =876
                    Height =276
                    FontSize =10
                    TabIndex =17
                    Name ="Residual"
                    ControlSource ="Residual"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="New Residual value for this unit"
                    GroupTable =1

                    LayoutCachedLeft =20184
                    LayoutCachedTop =36
                    LayoutCachedWidth =21060
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =21132
                    Top =36
                    Width =36
                    Height =276
                    FontSize =9
                    TabIndex =18
                    BorderColor =15527148
                    Name ="OrigResidual"
                    ControlSource ="OrigResidual"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =21132
                    LayoutCachedTop =36
                    LayoutCachedWidth =21168
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =3096
                    Top =36
                    Width =1572
                    Height =276
                    FontSize =10
                    TabIndex =3
                    BackColor =16317438
                    Name ="OptionSubType"
                    ControlSource ="OptionSubType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ExitOptionTypes.OptionSubType, LookUp_ExitOptionTypes.OptionSubTyp"
                        "eDesc FROM LookUp_ExitOptionTypes WHERE (((LookUp_ExitOptionTypes.OptionTypeID)>"
                        "2 Or (LookUp_ExitOptionTypes.OptionTypeID)<0));"
                    ColumnWidths ="1728;2160"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3096
                    LayoutCachedTop =36
                    LayoutCachedWidth =4668
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12924
                    Top =36
                    Width =468
                    Height =276
                    FontSize =8
                    TabIndex =9
                    Name ="AddExtend"
                    Caption ="▫EXT"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add another Extend Option to this Unit"
                    GroupTable =1

                    LayoutCachedLeft =12924
                    LayoutCachedTop =36
                    LayoutCachedWidth =13392
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =0
                    PressedThemeColorIndex =0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =16980
                    Top =36
                    Width =300
                    Height =276
                    FontSize =10
                    TabIndex =15
                    ForeColor =1643706
                    Name ="DeleteMe"
                    Caption ="Õ"
                    OnClick ="[Event Procedure]"
                    FontName ="Wingdings"
                    ControlTipText ="Delete This Unit's Option"
                    GroupTable =1

                    LayoutCachedLeft =16980
                    LayoutCachedTop =36
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =1643706
                    PressedThemeColorIndex =9
                    PressedShade =75.0
                    PressedForeColor =1643706
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14004
                    Top =36
                    Width =480
                    Height =276
                    FontSize =8
                    TabIndex =11
                    Name ="ModifyTERM"
                    Caption ="TERM"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Terminate Option for this unit"
                    GroupTable =1

                    LayoutCachedLeft =14004
                    LayoutCachedTop =36
                    LayoutCachedWidth =14484
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =790406
                    BorderColor =15527148
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =790406
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13464
                    Top =36
                    Width =468
                    Height =276
                    FontSize =8
                    TabIndex =10
                    Name ="ModifyExch"
                    Caption ="EXCH"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set ExchangeIT Option for this unit"
                    GroupTable =1

                    LayoutCachedLeft =13464
                    LayoutCachedTop =36
                    LayoutCachedWidth =13932
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8872517
                    BorderColor =15527148
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17340
                    Top =36
                    Width =2772
                    Height =276
                    FontSize =9
                    TabIndex =16
                    Name ="EarlyGapOverLapMsg"
                    ControlSource ="[EarlyGapOverLapMsg]"
                    ControlTipText ="Expiry and or Overlap messages will appear here, as needed."
                    GroupTable =1

                    LayoutCachedLeft =17340
                    LayoutCachedTop =36
                    LayoutCachedWidth =20112
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =19
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14544
                    Top =36
                    Width =12
                    Height =276
                    TabIndex =12
                    Name ="LeaseOptionID"
                    ControlSource ="LeaseOptionID"
                    GroupTable =1

                    LayoutCachedLeft =14544
                    LayoutCachedTop =36
                    LayoutCachedWidth =14556
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1872
                    Top =36
                    Width =1152
                    Height =276
                    FontSize =10
                    TabIndex =1
                    Name ="OptionTypeDesc"
                    ControlSource ="OptionTypeDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ExitOptionTypes.ExitOptionType FROM LookUp_ExitOptionTypes WHERE ("
                        "((LookUp_ExitOptionTypes.OptionTypeID)>2 Or (LookUp_ExitOptionTypes.OptionTypeID"
                        ")<0)) GROUP BY LookUp_ExitOptionTypes.ExitOptionType;"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Extended Fee for this unit"
                    GroupTable =1

                    LayoutCachedLeft =1872
                    LayoutCachedTop =36
                    LayoutCachedWidth =3024
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =14265
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =14265
                    LayoutCachedWidth =14294
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14115
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =14115
                    LayoutCachedWidth =14144
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =150
                    Width =765
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =2
                    Name ="txtCurrent"

                    LayoutCachedLeft =150
                    LayoutCachedWidth =915
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13740
                    Width =29
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    ForeColor =12566463
                    Name ="GroupID"
                    ControlSource ="GroupID"

                    LayoutCachedLeft =13740
                    LayoutCachedWidth =13769
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
            End
        End
    End
End
CodeBehindForm
' See "ExitOptions_subform_Extend.cls"
