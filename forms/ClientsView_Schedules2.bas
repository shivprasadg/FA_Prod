Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16200
    DatasheetFontHeight =10
    ItemSuffix =538
    Left =8220
    Top =3870
    Right =20445
    Bottom =9990
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x90e8da168ba3e440
    End
    DatasheetFontName ="Arial"
    OnActivate ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Transactions"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
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
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =669
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8355
                    Top =45
                    Width =1605
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label124"
                    Caption ="┌── Schedule ──┐"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =8355
                    LayoutCachedTop =45
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =285
                    ColumnStart =5
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =95
                    Top =270
                    Width =16155
                    Height =300
                    BorderColor =-2147483615
                    Name ="Box130"
                    HorizontalAnchor =2
                    LayoutCachedTop =270
                    LayoutCachedWidth =16155
                    LayoutCachedHeight =570
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =30
                    Width =1005
                    Height =195
                    ColumnOrder =0
                    FontSize =9
                    Name ="Text84"
                    ControlSource ="SchID"
                    FontName ="Calibri"

                    LayoutCachedLeft =14160
                    LayoutCachedTop =30
                    LayoutCachedWidth =15165
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15225
                    Top =30
                    Width =795
                    Height =195
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =1
                    Name ="MLID"
                    ControlSource ="MLID"
                    FontName ="Calibri"

                    LayoutCachedLeft =15225
                    LayoutCachedTop =30
                    LayoutCachedWidth =16020
                    LayoutCachedHeight =225
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =3255
                    Top =285
                    Width =2970
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label20"
                    Caption ="Description"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3255
                    LayoutCachedTop =285
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8355
                    Top =285
                    Width =765
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="Label42"
                    Caption =" Sent"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =8355
                    LayoutCachedTop =285
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9195
                    Top =285
                    Width =765
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="Label43"
                    Caption ="Back"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =9195
                    LayoutCachedTop =285
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =990
                    Top =285
                    Width =2205
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label46"
                    Caption ="Schedule "
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =990
                    LayoutCachedTop =285
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =10035
                    Top =285
                    Width =1470
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label63"
                    Caption ="Assignee"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10035
                    LayoutCachedTop =285
                    LayoutCachedWidth =11505
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7485
                    Top =285
                    Width =810
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label70"
                    Caption ="Last Del"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7485
                    LayoutCachedTop =285
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =45
                    Top =285
                    Width =885
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label81"
                    Caption ="FAID"
                    FontName ="Calibri"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =45
                    LayoutCachedTop =285
                    LayoutCachedWidth =930
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6285
                    Top =285
                    Width =1140
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label101"
                    Caption ="OEC"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6285
                    LayoutCachedTop =285
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11565
                    Top =285
                    Width =195
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =16777215
                    Name ="Label102"
                    Caption ="I"
                    FontName ="Calibri"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11565
                    LayoutCachedTop =285
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =45
                    Top =45
                    Width =885
                    Name ="EmptyCell136"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =45
                    LayoutCachedTop =45
                    LayoutCachedWidth =930
                    LayoutCachedHeight =285
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =990
                    Top =45
                    Width =2205
                    Name ="EmptyCell137"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =990
                    LayoutCachedTop =45
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =285
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3255
                    Top =45
                    Width =2970
                    Name ="EmptyCell141"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3255
                    LayoutCachedTop =45
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =285
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6285
                    Top =45
                    Width =1140
                    Name ="EmptyCell142"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6285
                    LayoutCachedTop =45
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =285
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7485
                    Top =45
                    Width =810
                    Name ="EmptyCell143"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7485
                    LayoutCachedTop =45
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =285
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10035
                    Top =45
                    Width =1470
                    Name ="EmptyCell150"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10035
                    LayoutCachedTop =45
                    LayoutCachedWidth =11505
                    LayoutCachedHeight =285
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11565
                    Top =45
                    Width =195
                    Name ="EmptyCell280"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11565
                    LayoutCachedTop =45
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =285
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =11775
                    Top =285
                    Width =1380
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =8210719
                    BorderColor =8210719
                    ForeColor =16777215
                    Name ="Schedule Label"
                    Caption ="Schedule Reports                    "
                    FontName ="Calibri"
                    EventProcPrefix ="Schedule_Label"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11775
                    LayoutCachedTop =285
                    LayoutCachedWidth =13155
                    LayoutCachedHeight =525
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =3
                    PictureAlignment =0
                    PictureType =2
                    Top =525
                    Width =16155
                    Height =144
                    Name ="Image142"
                    Picture ="DkBlue_Shadow"
                    TopPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =525
                    LayoutCachedWidth =16155
                    LayoutCachedHeight =669
                    TabIndex =2
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =9
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =14085
                    Top =270
                    Width =2070
                    Height =240
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =8210719
                    BorderColor =8210719
                    ForeColor =16777215
                    Name ="Label537"
                    Caption ="Exchanged By Sch-Grp         "
                    FontName ="Calibri"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14085
                    LayoutCachedTop =270
                    LayoutCachedWidth =16155
                    LayoutCachedHeight =510
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =382
            Name ="Detail"
            OnClick ="[Event Procedure]"
            AlternateBackColor =16249583
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8085
                    Top =330
                    Width =390
                    Height =30
                    TabIndex =12
                    ForeColor =-2147483617
                    Name ="Client"
                    ControlSource ="Client"
                    StatusBarText ="link to client file"
                    FontName ="Calibri"

                    LayoutCachedLeft =8085
                    LayoutCachedTop =330
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8565
                    Top =330
                    Width =390
                    Height =30
                    TabIndex =13
                    ForeColor =-2147483617
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    StatusBarText ="Short Name for Reporting"
                    FontName ="Calibri"

                    LayoutCachedLeft =8565
                    LayoutCachedTop =330
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9105
                    Top =330
                    Width =465
                    Height =30
                    TabIndex =15
                    ForeColor =-2147483617
                    Name ="SchID"
                    ControlSource ="SchID"
                    FontName ="Calibri"

                    LayoutCachedLeft =9105
                    LayoutCachedTop =330
                    LayoutCachedWidth =9570
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3255
                    Top =75
                    Width =2970
                    Height =285
                    FontSize =9
                    TabIndex =2
                    Name ="dealShortDesc"
                    ControlSource ="SchShortDesc"
                    Format =">"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =3255
                    LayoutCachedTop =75
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8355
                    Top =75
                    Width =765
                    Height =285
                    FontSize =9
                    TabIndex =5
                    Name ="docSchedSent"
                    ControlSource ="SchOut"
                    Format ="mm/dd/yy"
                    FontName ="Calibri"
                    InputMask ="\"Sched Sent\""
                    GroupTable =1
                    RightPadding =38

                    LayoutCachedLeft =8355
                    LayoutCachedTop =75
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9195
                    Top =75
                    Width =765
                    Height =285
                    FontSize =9
                    TabIndex =6
                    Name ="docSchedBack"
                    ControlSource ="SchIn"
                    Format ="mm/dd/yy"
                    FontName ="Calibri"
                    GroupTable =1
                    RightPadding =38

                    LayoutCachedLeft =9195
                    LayoutCachedTop =75
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =330
                    Width =1065
                    Height =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    Format =">"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    FontName ="Calibri"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =330
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =330
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =990
                    Top =75
                    Width =2205
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =12349952
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click To Open Schedule"
                    GroupTable =1

                    LayoutCachedLeft =990
                    LayoutCachedTop =75
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6285
                    Top =75
                    Width =1140
                    Height =285
                    FontSize =9
                    TabIndex =3
                    Name ="SumOfOrigEquipCost"
                    ControlSource ="ECost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =6285
                    LayoutCachedTop =75
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5595
                    Top =330
                    Width =765
                    Height =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    Format =">"
                    FontName ="Calibri"

                    LayoutCachedLeft =5595
                    LayoutCachedTop =330
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =330
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10035
                    Top =75
                    Width =1470
                    Height =285
                    FontSize =9
                    TabIndex =7
                    Name ="Text65"
                    ControlSource ="Bassign"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =10035
                    LayoutCachedTop =75
                    LayoutCachedWidth =11505
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7485
                    Top =75
                    Width =810
                    Height =285
                    FontSize =9
                    TabIndex =4
                    Name ="MaxOfUnitEstDelDate"
                    ControlSource ="MaxOfUnitEstDelDate"
                    Format ="mm/yy"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =7485
                    LayoutCachedTop =75
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =45
                    Top =75
                    Width =885
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="SchFAID"
                    ControlSource ="SchFAID"
                    StatusBarText ="deal id"
                    FontName ="Calibri"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    GroupTable =1

                    LayoutCachedLeft =45
                    LayoutCachedTop =75
                    LayoutCachedWidth =930
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4740
                    Top =330
                    Width =555
                    Height =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =10
                    ForeColor =12349952
                    Name ="SCHNo1"
                    ControlSource ="SCHNo1"
                    Format ="Fixed"
                    StatusBarText ="added 8/15/2013 schedule number for sequencing"
                    FontName ="Calibri"
                    ControlTipText ="Double click to open form to reset schedul order"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="SCHORDf"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"SCHNo1\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">SCHORDf</Argument"
                        End
                        Begin
                            Comment ="_AXL:></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4740
                    LayoutCachedTop =330
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =330
                    RowStart =2
                    RowEnd =2
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11565
                    Top =75
                    Width =195
                    Height =285
                    TabIndex =8
                    Name ="Assigned_Indicator\015\012\015\012\015\012\015\012"
                    ControlSource ="=DMax(\"Assigned_Indicator\",\"SchGrp\",\"SGrpSchid=\" & [Schid])"
                    FontName ="Calibri"
                    EventProcPrefix ="Assigned_Indicator________"
                    GroupTable =1

                    LayoutCachedLeft =11565
                    LayoutCachedTop =75
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =11820
                    Top =75
                    Width =2220
                    Height =285
                    TabIndex =14
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags, LaunchReportsList.FormGroup FROM LaunchReportsList WH"
                        "ERE (((LaunchReportsList.SecurityLevel)<=9) AND ((LaunchReportsList.FormGroup)=\""
                        "Schedule\")) ORDER BY LaunchReportsList.ListOrder;"
                    ColumnWidths ="2016;0;0;0"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Docs Report (Excel)"
                    AllowValueListEdits =1
                    InheritValueList =1

                    LayoutCachedLeft =11820
                    LayoutCachedTop =75
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14100
                    Top =60
                    Width =1980
                    Height =285
                    FontSize =9
                    TabIndex =16
                    Name ="Text536"
                    ControlSource ="NewSchGrp"
                    FontName ="Calibri"
                    GroupTable =2

                    LayoutCachedLeft =14100
                    LayoutCachedTop =60
                    LayoutCachedWidth =16080
                    LayoutCachedHeight =345
                    LayoutGroup =2
                    GroupTable =2
                End
            End
        End
        Begin FormFooter
            Height =15
            BackColor =16249583
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "ClientsView_Schedules2.cls"
