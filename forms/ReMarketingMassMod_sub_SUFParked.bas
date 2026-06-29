Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20970
    DatasheetFontHeight =11
    ItemSuffix =556
    Left =13110
    Top =10170
    Right =-30916
    Bottom =23250
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6272a400b67fe540
    End
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
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1302
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =600
                    Width =20189
                    Height =285
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =20189
                    LayoutCachedHeight =885
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =900
                    Width =20189
                    Height =396
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =900
                    LayoutCachedWidth =20189
                    LayoutCachedHeight =1296
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4800
                    Top =60
                    Width =6135
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label88"
                    Caption =" ┌────SUF Dates────┐              ┌─────────  Parked ───────────────────────┐"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =60
                    LayoutCachedWidth =10935
                    LayoutCachedHeight =330
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =14040
                    Top =60
                    Width =1728
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label417"
                    Caption =" ┌──── Off Lease ──┐"
                    LayoutCachedLeft =14040
                    LayoutCachedTop =60
                    LayoutCachedWidth =15768
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =2
                End
                Begin EmptyCell
                    Left =36
                    Top =300
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =300
                    LayoutCachedWidth =36
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =300
                    Width =828
                    Name ="EmptyCell97"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =300
                    LayoutCachedWidth =924
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =996
                    Top =300
                    Width =1188
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="send2Excel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =300
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =540
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
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
                    Left =2244
                    Top =300
                    Width =2100
                    Name ="EmptyCell99"
                    GroupTable =1
                    LayoutCachedLeft =2244
                    LayoutCachedTop =300
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =540
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =4416
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="SUFSentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4416
                    LayoutCachedTop =300
                    LayoutCachedWidth =5568
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26265
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26265
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5640
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="SUFReturnedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5640
                    LayoutCachedTop =300
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =540
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26265
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26265
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6864
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="ParkedDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6864
                    LayoutCachedTop =300
                    LayoutCachedWidth =8016
                    LayoutCachedHeight =540
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26316
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =24768
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8088
                    Top =300
                    Width =2808
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="LocParkedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8088
                    LayoutCachedTop =300
                    LayoutCachedWidth =10896
                    LayoutCachedHeight =540
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26316
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26316
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =10968
                    Top =300
                    Width =960
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="MileageSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10968
                    LayoutCachedTop =300
                    LayoutCachedWidth =11928
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26316
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26316
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12000
                    Top =300
                    Width =840
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="HoursSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12000
                    LayoutCachedTop =300
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26316
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26316
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12912
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="BankNotifiedSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12912
                    LayoutCachedTop =300
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =6108695
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14136
                    Top =300
                    Width =648
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="UnitStatusSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14136
                    LayoutCachedTop =300
                    LayoutCachedWidth =14784
                    LayoutCachedHeight =540
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14856
                    Top =300
                    Width =1152
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="OffLeaseSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14856
                    LayoutCachedTop =300
                    LayoutCachedWidth =16008
                    LayoutCachedHeight =540
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =36
                    Top =600
                    Width =0
                    Height =315
                    Name ="EmptyCell68"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =600
                    LayoutCachedWidth =36
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =600
                    Width =828
                    Height =315
                    Name ="EmptyCell69"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =600
                    LayoutCachedWidth =924
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =996
                    Top =600
                    Width =1188
                    Height =315
                    Name ="EmptyCell70"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =600
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2244
                    Top =600
                    Width =2100
                    Height =315
                    Name ="EmptyCell71"
                    GroupTable =1
                    LayoutCachedLeft =2244
                    LayoutCachedTop =600
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4416
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Sent"
                    GroupTable =1
                    LayoutCachedLeft =4416
                    LayoutCachedTop =600
                    LayoutCachedWidth =5568
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5640
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Returned"
                    GroupTable =1
                    LayoutCachedLeft =5640
                    LayoutCachedTop =600
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6864
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Park Date"
                    GroupTable =1
                    LayoutCachedLeft =6864
                    LayoutCachedTop =600
                    LayoutCachedWidth =8016
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8088
                    Top =600
                    Width =2808
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label544"
                    Caption ="Parked Location"
                    GroupTable =1
                    LayoutCachedLeft =8088
                    LayoutCachedTop =600
                    LayoutCachedWidth =10896
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10968
                    Top =600
                    Width =960
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption =" Mileage"
                    GroupTable =1
                    LayoutCachedLeft =10968
                    LayoutCachedTop =600
                    LayoutCachedWidth =11928
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12000
                    Top =600
                    Width =840
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption ="Hours"
                    GroupTable =1
                    LayoutCachedLeft =12000
                    LayoutCachedTop =600
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12912
                    Top =600
                    Width =1152
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Bank Noti'ed"
                    GroupTable =1
                    LayoutCachedLeft =12912
                    LayoutCachedTop =600
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =14136
                    Top =600
                    Width =648
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label379"
                    Caption ="Status"
                    GroupTable =1
                    LayoutCachedLeft =14136
                    LayoutCachedTop =600
                    LayoutCachedWidth =14784
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =14856
                    Top =600
                    Width =1152
                    Height =315
                    ForeColor =16777215
                    Name ="Label41"
                    Caption ="Off Lease "
                    GroupTable =1
                    LayoutCachedLeft =14856
                    LayoutCachedTop =600
                    LayoutCachedWidth =16008
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =36
                    Top =984
                    Width =0
                    Height =288
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =984
                    LayoutCachedWidth =36
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =96
                    Top =984
                    Width =828
                    Height =288
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =10040879
                    Name ="Label20"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =984
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =996
                    Top =984
                    Width =1188
                    Height =288
                    Name ="Label1"
                    Caption ="Unit Num"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =984
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4416
                    Top =984
                    Width =1152
                    Height =288
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =12
                    Name ="SufSentCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =4416
                    LayoutCachedTop =984
                    LayoutCachedWidth =5568
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =2244
                            Top =984
                            Width =2100
                            Height =288
                            Name ="Label2"
                            Caption ="VIN"
                            GroupTable =1
                            LayoutCachedLeft =2244
                            LayoutCachedTop =984
                            LayoutCachedWidth =4344
                            LayoutCachedHeight =1272
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Top =984
                    Width =1152
                    Height =288
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =13
                    Name ="SUFReturnedCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =5640
                    LayoutCachedTop =984
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6864
                    Top =984
                    Width =1152
                    Height =288
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =14
                    Name ="ParkedDateCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =6864
                    LayoutCachedTop =984
                    LayoutCachedWidth =8016
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =8088
                    Top =984
                    Width =2808
                    Height =288
                    FontSize =9
                    TabIndex =15
                    Name ="LocParkedCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)<>\"billing\") AND ((Locations.ClientGroupID"
                        ")=771) AND ((Locations.LocObsolete)=No)) ORDER BY Locations.LocationName, Locati"
                        "ons.LState;"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =8088
                    LayoutCachedTop =984
                    LayoutCachedWidth =10896
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10968
                    Top =984
                    Width =960
                    Height =288
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =16
                    Name ="MileageCBO"
                    Format ="Standard"
                    ControlTipText ="This Value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =10968
                    LayoutCachedTop =984
                    LayoutCachedWidth =11928
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12000
                    Top =984
                    Width =840
                    Height =288
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =17
                    Name ="HoursCbo"
                    Format ="Standard"
                    Tag ="Medit"
                    ControlTipText ="This value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =12000
                    LayoutCachedTop =984
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12912
                    Top =984
                    Width =1152
                    Height =288
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =18
                    Name ="BankNotifiedCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =12912
                    LayoutCachedTop =984
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14136
                    Top =984
                    Width =648
                    Height =288
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =19
                    BorderColor =9211020
                    Name ="UnitStatusCBO"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =14136
                    LayoutCachedTop =984
                    LayoutCachedWidth =14784
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14856
                    Top =984
                    Width =1152
                    Height =288
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =20
                    BorderColor =9211020
                    Name ="OffLeaseCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1

                    LayoutCachedLeft =14856
                    LayoutCachedTop =984
                    LayoutCachedWidth =16008
                    LayoutCachedHeight =1272
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15945
                    Top =300
                    Width =4245
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="UnitNoteSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Notes below to all units listed"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15945
                    LayoutCachedTop =300
                    LayoutCachedWidth =20190
                    LayoutCachedHeight =540
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =15945
                    Top =600
                    Width =4245
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label526"
                    Caption ="Add a Note to the Unit"
                    LayoutCachedLeft =15945
                    LayoutCachedTop =600
                    LayoutCachedWidth =20190
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15945
                    Top =975
                    Width =4215
                    Height =285
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =21
                    BorderColor =9211020
                    Name ="UnitNoteCBO"
                    Tag ="Medit"
                    ControlTipText ="The text here will be INSERT into the Unit's Notes as a NEW Note"

                    LayoutCachedLeft =15945
                    LayoutCachedTop =975
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1260
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20910
                    Width =60
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="Text548"
                    ControlSource ="ClientGroupID"
                    ControlTipText ="Date we were invoiced for the Inspection"

                    LayoutCachedLeft =20910
                    LayoutCachedWidth =20970
                    LayoutCachedHeight =255
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15945
                    Top =30
                    Width =4245
                    Height =255
                    FontSize =10
                    TabIndex =13
                    BorderColor =9211020
                    Name ="UnitNote"
                    Tag ="Medit"
                    ControlTipText ="The text here will be INSERT into the Unit's Notes as a NEW Note"

                    LayoutCachedLeft =15945
                    LayoutCachedTop =30
                    LayoutCachedWidth =20190
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =20250
                    Top =30
                    Width =660
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    Name ="SaveThisNote"
                    Caption ="◄Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20250
                    LayoutCachedTop =30
                    LayoutCachedWidth =20910
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26316
                    BorderColor =6108695
                    HoverColor =5676533
                    PressedColor =26316
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =0
                    Height =252
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =36
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =96
                    Top =36
                    Width =828
                    Height =252
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =10040879
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =96
                    LayoutCachedTop =36
                    LayoutCachedWidth =924
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =36
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2244
                    Top =36
                    Width =2100
                    Height =252
                    FontSize =10
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =2244
                    LayoutCachedTop =36
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4416
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =4
                    BackColor =15070975
                    Name ="SufSent"
                    ControlSource ="SUF Sent"
                    Format ="Short Date"
                    ValidationText ="BO Date is required"
                    ControlTipText ="Date the SUF was sent to client"
                    GroupTable =1

                    LayoutCachedLeft =4416
                    LayoutCachedTop =36
                    LayoutCachedWidth =5568
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5640
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =5
                    BackColor =15070975
                    Name ="SUFReturnedf"
                    ControlSource ="SUF Returned"
                    Format ="Short Date"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Date the SUF was returned back from the client"
                    GroupTable =1

                    LayoutCachedLeft =5640
                    LayoutCachedTop =36
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6864
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =6
                    BackColor =15070975
                    Name ="ParkedDate"
                    ControlSource ="ParkDT"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date the SUF states the unit was parked"
                    GroupTable =1

                    LayoutCachedLeft =6864
                    LayoutCachedTop =36
                    LayoutCachedWidth =8016
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =8088
                    Top =36
                    Width =2808
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="LocationParkedID"
                    ControlSource ="LocationParkedID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)<>\"billing\") AND ((Locations.ClientGroupID"
                        ")=771) AND ((Locations.LocObsolete)=No)) ORDER BY Locations.LocationName, Locati"
                        "ons.LState;"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    GroupTable =1

                    LayoutCachedLeft =8088
                    LayoutCachedTop =36
                    LayoutCachedWidth =10896
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10968
                    Top =36
                    Width =960
                    Height =252
                    FontSize =10
                    TabIndex =8
                    BackColor =15070975
                    Name ="MiParked"
                    ControlSource ="MiParked"
                    Format ="Standard"
                    DefaultValue ="0"
                    ControlTipText ="Mileage the SUF states when the unit was parked"
                    GroupTable =1

                    LayoutCachedLeft =10968
                    LayoutCachedTop =36
                    LayoutCachedWidth =11928
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12000
                    Top =36
                    Width =840
                    Height =252
                    FontSize =10
                    TabIndex =9
                    BackColor =15070975
                    Name ="Hours"
                    ControlSource ="HrParked"
                    Format ="Standard"
                    ControlTipText ="Hours the SUF states when the unit was parked"
                    GroupTable =1

                    LayoutCachedLeft =12000
                    LayoutCachedTop =36
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12912
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =10
                    BackColor =16317438
                    Name ="BankNotified"
                    ControlSource ="BankNotified"
                    Format ="Short Date"
                    ControlTipText ="Date the Assignee Bank was notified that unit was parked"
                    GroupTable =1

                    LayoutCachedLeft =12912
                    LayoutCachedTop =36
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14136
                    Top =36
                    Width =648
                    Height =252
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    ControlTipText ="Date we were invoiced for the Inspection"
                    GroupTable =1

                    LayoutCachedLeft =14136
                    LayoutCachedTop =36
                    LayoutCachedWidth =14784
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    FontItalic = NotDefault
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14856
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="OffLease"
                    ControlSource ="unitoffleasedt"
                    Format ="Short Date"
                    ControlTipText ="Amount we were invoiced for the Inspection"
                    GroupTable =1

                    LayoutCachedLeft =14856
                    LayoutCachedTop =36
                    LayoutCachedWidth =16008
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =315
                    Width =20160
                    Height =59
                    FontSize =10
                    TabIndex =15
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
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9810
                    Width =29
                    Height =270
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =2
                    Name ="SchNo"
                    ControlSource ="SchNo"

                    LayoutCachedLeft =9810
                    LayoutCachedWidth =9839
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =9690
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =9690
                    LayoutCachedWidth =9719
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9540
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =9540
                    LayoutCachedWidth =9569
                    LayoutCachedHeight =270
                End
            End
        End
    End
End
CodeBehindForm
' See "ReMarketingMassMod_sub_SUFParked.cls"
