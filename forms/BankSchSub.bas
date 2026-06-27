Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18540
    DatasheetFontHeight =10
    ItemSuffix =95
    Left =8760
    Top =8370
    Right =26160
    Bottom =15810
    ShortcutMenuBar ="SortClip"
    OrderBy ="ClientShNm, BLCD DESC"
    RecSrcDt = Begin
        0x1585e6218ca3e440
    End
    RecordSource ="BankSchSummq"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Transactions"
    ShowPageMargins =0
    AllowLayoutView =0
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
            Height =532
            BackColor =15527148
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4725
                    Top =60
                    Width =3630
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label20"
                    Caption ="Description"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =4725
                    LayoutCachedTop =60
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =510
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =780
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label45"
                    Caption ="Master No"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =510
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2220
                    Top =60
                    Width =1620
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label46"
                    Caption ="Sch No"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =510
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =15045
                    Top =60
                    Width =1500
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label48"
                    Caption ="Equipment Cost"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =15045
                    LayoutCachedTop =60
                    LayoutCachedWidth =16545
                    LayoutCachedHeight =510
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8415
                    Top =60
                    Width =1080
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label70"
                    Caption ="Comm \015\012Date"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =8415
                    LayoutCachedTop =60
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =510
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9555
                    Top =60
                    Width =660
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label72"
                    Caption ="Primary Term"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =9555
                    LayoutCachedTop =60
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =510
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =13545
                    Top =60
                    Width =1440
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label73"
                    Caption ="Rent"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =13545
                    LayoutCachedTop =60
                    LayoutCachedWidth =14985
                    LayoutCachedHeight =510
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =10275
                    Top =60
                    Width =1200
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label74"
                    Caption ="1st Rent Assigned \015\012"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =10275
                    LayoutCachedTop =60
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =510
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =3900
                    Top =60
                    Width =765
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label76"
                    Caption ="Grp"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =510
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11535
                    Top =60
                    Width =945
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label79"
                    Caption ="Expiry"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =11535
                    LayoutCachedTop =60
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =510
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =12540
                    Top =60
                    Width =945
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label81"
                    Caption ="ExIT"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =12540
                    LayoutCachedTop =60
                    LayoutCachedWidth =13485
                    LayoutCachedHeight =510
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =480
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label82"
                    Caption =" "
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =480
                    LayoutCachedHeight =510
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =660
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label84"
                    Caption ="Assignee_Bank\015\012\015\012\015\012\015\012"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =660
                    LayoutCachedHeight =510
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =720
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label85"
                    Caption =" "
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =720
                    LayoutCachedHeight =510
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label86"
                    Caption =" "
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =600
                    LayoutCachedTop =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =510
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =540
                    Top =60
                    Width =0
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483611
                    Name ="Label87"
                    Caption ="SchID"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =540
                    LayoutCachedTop =60
                    LayoutCachedWidth =540
                    LayoutCachedHeight =510
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =840
                    Top =60
                    Width =1320
                    Height =450
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Client"
                    FontName ="Calibri"
                    GroupTable =1
                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =510
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =315
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4725
                    Top =45
                    Width =3630
                    Height =210
                    FontSize =9
                    TabIndex =10
                    Name ="dealShortDesc"
                    ControlSource ="SchShortDesc"
                    Format =">"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =4725
                    LayoutCachedTop =45
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =45
                    Width =0
                    Height =210
                    FontSize =9
                    TabIndex =6
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    Format =">"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =780
                    LayoutCachedTop =45
                    LayoutCachedWidth =780
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =45
                    Width =1620
                    Height =210
                    FontSize =9
                    TabIndex =8
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    StatusBarText ="schedule number starts with 1"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =2220
                    LayoutCachedTop =45
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15045
                    Top =45
                    Width =1500
                    Height =210
                    FontSize =9
                    TabIndex =17
                    BackColor =15858167
                    Name ="SumOfOrigEquipCost"
                    ControlSource ="SumOfOrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =15045
                    LayoutCachedTop =45
                    LayoutCachedWidth =16545
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =45
                    Width =1320
                    Height =210
                    FontSize =9
                    TabIndex =7
                    Name ="Client"
                    ControlSource ="ClientShNm"
                    StatusBarText ="link to client file"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =840
                    LayoutCachedTop =45
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =45
                    Width =0
                    Height =210
                    TabIndex =3
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    StatusBarText ="Short Name for Reporting"
                    GroupTable =1

                    LayoutCachedLeft =600
                    LayoutCachedTop =45
                    LayoutCachedWidth =600
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =480
                    Top =45
                    Width =0
                    Height =210
                    TabIndex =1
                    Name ="MLID"
                    ControlSource ="MLID"
                    GroupTable =1

                    LayoutCachedLeft =480
                    LayoutCachedTop =45
                    LayoutCachedWidth =480
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8415
                    Top =45
                    Width =1080
                    Height =210
                    FontSize =9
                    TabIndex =11
                    Name ="BLCD\015\012\015\012\015\012\015\012"
                    ControlSource ="BLCD"
                    Format ="Short Date"
                    FontName ="Calibri"
                    EventProcPrefix ="BLCD________"
                    GroupTable =1

                    LayoutCachedLeft =8415
                    LayoutCachedTop =45
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9555
                    Top =45
                    Width =660
                    Height =210
                    ColumnWidth =1455
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    ForeColor =-2147483615
                    Name ="LeaseTermPrimary"
                    ControlSource ="LeaseTermPrimary"
                    Format ="Fixed"
                    StatusBarText ="primary term in months"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =9555
                    LayoutCachedTop =45
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13545
                    Top =45
                    Height =210
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    ForeColor =4754549
                    Name ="LeasePymtForGroup"
                    ControlSource ="LeasePymtForGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =13545
                    LayoutCachedTop =45
                    LayoutCachedWidth =14985
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10275
                    Top =45
                    Width =1200
                    Height =210
                    FontSize =9
                    TabIndex =13
                    Name ="SchAssDt"
                    ControlSource ="AssignmentDate"
                    Format ="Short Date"
                    StatusBarText ="if ML other, date of schedule outright assignment to FA"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =10275
                    LayoutCachedTop =45
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3900
                    Top =45
                    Width =765
                    Height =210
                    FontSize =9
                    TabIndex =9
                    Name ="SGrpNO"
                    ControlSource ="SGrpNO"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =3900
                    LayoutCachedTop =45
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =45
                    Width =0
                    Height =210
                    TabIndex =5
                    Name ="clientID"
                    ControlSource ="Client"
                    GroupTable =1

                    LayoutCachedLeft =720
                    LayoutCachedTop =45
                    LayoutCachedWidth =720
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =45
                    Width =0
                    Height =210
                    TabIndex =4
                    Name ="Assignee_Bank\015\012\015\012\015\012\015\012"
                    ControlSource ="Assignee_Bank"
                    StatusBarText ="select from Clients clientType=Bank"
                    EventProcPrefix ="Assignee_Bank________"
                    GroupTable =1

                    LayoutCachedLeft =660
                    LayoutCachedTop =45
                    LayoutCachedWidth =660
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11535
                    Top =45
                    Width =945
                    Height =210
                    FontSize =9
                    TabIndex =14
                    Name ="Exp"
                    ControlSource ="Exp"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =11535
                    LayoutCachedTop =45
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =45
                    Width =945
                    Height =210
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    BackColor =15592953
                    ForeColor =-2147483615
                    Name ="ExIT"
                    ControlSource ="ExIT"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GroupTable =1

                    LayoutCachedLeft =12540
                    LayoutCachedTop =45
                    LayoutCachedWidth =13485
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Top =45
                    Width =0
                    Height =210
                    TabIndex =2
                    Name ="SchID"
                    ControlSource ="SchID"
                    GroupTable =1

                    LayoutCachedLeft =540
                    LayoutCachedTop =45
                    LayoutCachedWidth =540
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =60
                    Top =30
                    Width =675
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="BnkSchBtn"
                    Caption ="Sched"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =735
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =16249583
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BankSchSub.cls"
