Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9803
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =7800
    Top =4425
    Right =18630
    Bottom =13470
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc7247113b6b5e540
    End
    RecordSource ="SELECT Schedule.SchID, Units.UnitUnitNum AS UnitNo, Units.UnitVin, Units.UnitAcc"
        "eptDate, Units.UnitDocAccpt, Units.UnitCOABack, Units.unitstatus, SchGrp.SGrpID,"
        " Units.UnitCofACopy FROM ((Schedule INNER JOIN SchGrp ON Schedule.SchID = SchGrp"
        ".SGrpSchID) INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF) INNER JOIN MstrL"
        "ease ON Schedule.SchMLRef = MstrLease.MLID WHERE (((MstrLease.MLOrig)=\"FA\")) O"
        "RDER BY Units.UnitUnitNum;"
    Caption ="CofAUnit List"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
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
        Begin ToggleButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =990
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4716
                    Top =660
                    Width =1392
                    Height =285
                    Name ="Label0"
                    Caption ="Unit#"
                    GroupTable =1
                    LayoutCachedLeft =4716
                    LayoutCachedTop =660
                    LayoutCachedWidth =6108
                    LayoutCachedHeight =945
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =36
                    Top =660
                    Width =2472
                    Height =285
                    Name ="Label1"
                    Caption ="VIN"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =660
                    LayoutCachedWidth =2508
                    LayoutCachedHeight =945
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6180
                    Top =660
                    Width =1308
                    Height =285
                    Name ="Label2"
                    Caption ="COA Back"
                    GroupTable =1
                    LayoutCachedLeft =6180
                    LayoutCachedTop =660
                    LayoutCachedWidth =7488
                    LayoutCachedHeight =945
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7560
                    Top =660
                    Width =756
                    Height =285
                    Name ="Label3"
                    Caption ="Status"
                    GroupTable =1
                    LayoutCachedLeft =7560
                    LayoutCachedTop =660
                    LayoutCachedWidth =8316
                    LayoutCachedHeight =945
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =3972
                    Top =660
                    Width =672
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label4"
                    Caption =" Acct"
                    GroupTable =1
                    LayoutCachedLeft =3972
                    LayoutCachedTop =660
                    LayoutCachedWidth =4644
                    LayoutCachedHeight =945
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2580
                    Top =660
                    Width =1332
                    Height =285
                    Name ="Label6"
                    Caption ="Accept Date"
                    GroupTable =1
                    LayoutCachedLeft =2580
                    LayoutCachedTop =660
                    LayoutCachedWidth =3912
                    LayoutCachedHeight =945
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4125
                    Top =225
                    Width =585
                    Height =405
                    FontSize =9
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Doc \015\012Only "
                    LayoutCachedLeft =4125
                    LayoutCachedTop =225
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =720
                    Height =315
                    TabIndex =1
                    ForeColor =10855845
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =600
                    Height =315
                    ForeColor =10855845
                    Name ="GroupID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8388
                    Top =660
                    Width =1320
                    Height =285
                    Name ="Label13"
                    Caption ="COA Copy"
                    GroupTable =1
                    LayoutCachedLeft =8388
                    LayoutCachedTop =660
                    LayoutCachedWidth =9708
                    LayoutCachedHeight =945
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8310
                    Top =330
                    Width =1320
                    Height =259
                    TabIndex =2
                    BorderColor =-2147483615
                    Name ="CofACopyCBO"
                    Format ="Short Date"

                    LayoutCachedLeft =8310
                    LayoutCachedTop =330
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =589
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2580
                    Top =330
                    Width =1320
                    Height =259
                    TabIndex =3
                    BorderColor =-2147483615
                    Name ="AcceptDateCBO"
                    Format ="Short Date"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =330
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =589
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =330
                    Width =1320
                    Height =259
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =-2147483615
                    Name ="COABacKCBO"
                    Format ="Short Date"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =330
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =589
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2580
                    Top =60
                    Width =1320
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="AcceptDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2580
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6120
                    Top =60
                    Width =1320
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="COABackSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8310
                    Top =60
                    Width =1320
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="CofACopySubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8310
                    LayoutCachedTop =60
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
            End
        End
        Begin Section
            Height =396
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4716
                    Top =48
                    Width =1392
                    Height =312
                    FontWeight =700
                    TabIndex =3
                    BackColor =15527148
                    BorderColor =-2147483615
                    ForeColor =8210719
                    Name ="UnitNo"
                    ControlSource ="UnitNo"
                    GroupTable =1

                    LayoutCachedLeft =4716
                    LayoutCachedTop =48
                    LayoutCachedWidth =6108
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =36
                    Top =48
                    Width =2472
                    Height =312
                    BorderColor =-2147483615
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =48
                    LayoutCachedWidth =2508
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6180
                    Top =48
                    Width =1308
                    Height =312
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =-2147483615
                    Name ="UnitCOABack"
                    ControlSource ="UnitCOABack"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =6180
                    LayoutCachedTop =48
                    LayoutCachedWidth =7488
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7560
                    Top =48
                    Width =756
                    Height =312
                    ColumnWidth =1965
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1

                    LayoutCachedLeft =7560
                    LayoutCachedTop =48
                    LayoutCachedWidth =8316
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2580
                    Top =48
                    Width =1332
                    Height =312
                    TabIndex =1
                    BorderColor =-2147483615
                    Name ="UnitAcceptDate"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =2580
                    LayoutCachedTop =48
                    LayoutCachedWidth =3912
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4224
                    Top =48
                    Width =420
                    Height =312
                    TabIndex =2
                    Name ="UnitDocAccpt"
                    ControlSource ="UnitDocAccpt"
                    GroupTable =1

                    LayoutCachedLeft =4224
                    LayoutCachedTop =48
                    LayoutCachedWidth =4644
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3972
                    Top =48
                    Width =192
                    Height =312
                    Name ="EmptyCell11"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =3972
                    LayoutCachedTop =48
                    LayoutCachedWidth =4164
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8388
                    Top =48
                    Width =1320
                    Height =312
                    TabIndex =6
                    BorderColor =-2147483615
                    Name ="UnitCofACopy"
                    ControlSource ="UnitCofACopy"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =8388
                    LayoutCachedTop =48
                    LayoutCachedWidth =9708
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
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
' See "Unit List_CofASub.cls"
