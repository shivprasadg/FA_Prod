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
    Width =15984
    DatasheetFontHeight =11
    ItemSuffix =771
    Left =29040
    Top =20235
    Right =-19816
    Bottom =28635
    ShortcutMenuBar ="FormsCmd"
    RecSrcDt = Begin
        0xc9fc7b31963be640
    End
    RecordSource ="SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Bi"
        "lldt, Billfrom, Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, "
        "AssigneeBankID, iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BC"
        "ontract  FROM (((Bills INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
        " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID) LEFT JOIN SchGrp ON"
        " Bills.GroupId = SchGrp.SGrpID) WHERE BillId IN(19335,19334,19336,19338,15134,15"
        "135,15136,15137)  ORDER BY Billdt DESC"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            Height =1332
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =600
                    Width =15984
                    Height =285
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =15984
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
                    Width =15984
                    Height =432
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =900
                    LayoutCachedWidth =15984
                    LayoutCachedHeight =1332
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =8
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =960
                    Width =1020
                    Height =518
                    FontSize =7
                    BackColor =14211288
                    BorderColor =14211288
                    Name ="Label63"
                    Caption ="Bills"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =518
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =9000
                    Top =60
                    Width =4515
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="   ┌───── Bill ─────┐                           ┌─────  Range ─────┐"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =60
                    LayoutCachedWidth =13515
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =180
                    Width =899
                    Height =300
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =4
                    Name ="BillCount"
                    ControlSource ="=Count([BillID])"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =180
                    LayoutCachedWidth =1919
                    LayoutCachedHeight =480
                    BackThemeColorIndex =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3300
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =1
                    Name ="InvoiceToExcel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Make a SINGLE Excel Export for the Selected Bill"
                    GroupTable =2
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3300
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =315
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =9
                    BackTint =40.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3795
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =2
                    Name ="Runbill"
                    Caption ="Run Bill"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Make a SINGLE Bill / Invoice for the Selected Bill"
                    GroupTable =2
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272f000000000727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffff0000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffff727272ff00000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffff00000000727272ae ,
                        0x727272ff727272ff727272ffffffffffffffffffffffffff727272ff727272ff ,
                        0x727272ff727272ae727272ffffffffffcdcdcdffa4a4a4ff00000000727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x00000000727272ff727272ffffffffffa3a3a3ffffffffff00000000727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffcdcdcdffa4a4a4ff00000000727272ff ,
                        0x727272ff00000000000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ffffffffffffffffffffffffff00000000727272b7 ,
                        0x727272ff00000000727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ff727272ba727272ffffffffffcdcdcdff9e9e9eff9e9e9e8100000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffff727272ff00000000 ,
                        0x0000000000000000727272ffffffffffa3a3a3ff9e9e9effa3a3a3ffffffffff ,
                        0xb7b7b7ff00000000727272ffffffffffffffffffffffffff727272ff00000000 ,
                        0x0000000000000000727272ffffffffffcdcdcdff9e9e9effcececeffffffffff ,
                        0xffffffff00000000727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ffffffffffd9d9d9ff787878f47272723c0000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ffd9d9d9ff787878f472727239000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff797979f37272723900000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272390000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3795
                    LayoutCachedTop =60
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =30
                    Top =300
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =300
                    LayoutCachedWidth =30
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =90
                    Top =300
                    Width =825
                    Name ="EmptyCell97"
                    GroupTable =1
                    LayoutCachedLeft =90
                    LayoutCachedTop =300
                    LayoutCachedWidth =915
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =975
                    Top =300
                    Width =1008
                    Name ="EmptyCell730"
                    GroupTable =1
                    LayoutCachedLeft =975
                    LayoutCachedTop =300
                    LayoutCachedWidth =1983
                    LayoutCachedHeight =540
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2040
                    Top =300
                    Width =576
                    Name ="EmptyCell735"
                    GroupTable =1
                    LayoutCachedLeft =2040
                    LayoutCachedTop =300
                    LayoutCachedWidth =2616
                    LayoutCachedHeight =540
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2670
                    Top =300
                    Width =1872
                    Name ="EmptyCell755"
                    GroupTable =1
                    LayoutCachedLeft =2670
                    LayoutCachedTop =300
                    LayoutCachedWidth =4542
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5910
                    Top =300
                    Width =1500
                    Name ="EmptyCell708"
                    GroupTable =1
                    LayoutCachedLeft =5910
                    LayoutCachedTop =300
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =540
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7470
                    Top =300
                    Width =1275
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="BillBankIDSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7470
                    LayoutCachedTop =300
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =540
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8805
                    Top =300
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="BilldtSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8805
                    LayoutCachedTop =300
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10020
                    Top =300
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="BillDueDtSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10020
                    LayoutCachedTop =300
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =11235
                    Top =300
                    Width =150
                    Name ="EmptyCell622"
                    GroupTable =1
                    LayoutCachedLeft =11235
                    LayoutCachedTop =300
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11445
                    Top =300
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="BillFromSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11445
                    LayoutCachedTop =300
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =540
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12660
                    Top =300
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="BilltoSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =300
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =540
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverColor =8965045
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13875
                    Top =300
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="BillSentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13875
                    LayoutCachedTop =300
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =540
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
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
                    Left =15090
                    Top =300
                    Width =855
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="DueonRecSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15090
                    LayoutCachedTop =300
                    LayoutCachedWidth =15945
                    LayoutCachedHeight =540
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =30
                    Top =600
                    Width =0
                    Height =315
                    Name ="EmptyCell68"
                    GroupTable =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =600
                    LayoutCachedWidth =30
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =90
                    Top =600
                    Width =825
                    Height =315
                    Name ="EmptyCell69"
                    GroupTable =1
                    LayoutCachedLeft =90
                    LayoutCachedTop =600
                    LayoutCachedWidth =915
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =975
                    Top =600
                    Width =1008
                    Height =315
                    Name ="EmptyCell731"
                    GroupTable =1
                    LayoutCachedLeft =975
                    LayoutCachedTop =600
                    LayoutCachedWidth =1983
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2040
                    Top =600
                    Width =576
                    Height =315
                    Name ="EmptyCell736"
                    GroupTable =1
                    LayoutCachedLeft =2040
                    LayoutCachedTop =600
                    LayoutCachedWidth =2616
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2670
                    Top =600
                    Width =1872
                    Height =315
                    Name ="EmptyCell756"
                    GroupTable =1
                    LayoutCachedLeft =2670
                    LayoutCachedTop =600
                    LayoutCachedWidth =4542
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4605
                    Top =600
                    Width =1245
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label629"
                    Caption =" "
                    GroupTable =1
                    LayoutCachedLeft =4605
                    LayoutCachedTop =600
                    LayoutCachedWidth =5850
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
                Begin EmptyCell
                    Left =5910
                    Top =600
                    Width =1500
                    Height =315
                    Name ="EmptyCell709"
                    GroupTable =1
                    LayoutCachedLeft =5910
                    LayoutCachedTop =600
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7470
                    Top =600
                    Width =1275
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label381"
                    Caption ="Bank"
                    GroupTable =1
                    LayoutCachedLeft =7470
                    LayoutCachedTop =600
                    LayoutCachedWidth =8745
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
                    Left =8805
                    Top =600
                    Width =1155
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Bill Date"
                    GroupTable =1
                    LayoutCachedLeft =8805
                    LayoutCachedTop =600
                    LayoutCachedWidth =9960
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
                    Left =10020
                    Top =600
                    Width =1155
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Bill Due"
                    GroupTable =1
                    LayoutCachedLeft =10020
                    LayoutCachedTop =600
                    LayoutCachedWidth =11175
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
                Begin EmptyCell
                    Left =11235
                    Top =600
                    Width =150
                    Height =315
                    Name ="EmptyCell623"
                    GroupTable =1
                    LayoutCachedLeft =11235
                    LayoutCachedTop =600
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =11445
                    Top =600
                    Width =1155
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Bill From"
                    GroupTable =1
                    LayoutCachedLeft =11445
                    LayoutCachedTop =600
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12660
                    Top =600
                    Width =1155
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Bill To"
                    GroupTable =1
                    LayoutCachedLeft =12660
                    LayoutCachedTop =600
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =13875
                    Top =600
                    Width =1155
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label383"
                    Caption ="Bill Sent"
                    GroupTable =1
                    LayoutCachedLeft =13875
                    LayoutCachedTop =600
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =15090
                    Top =600
                    Width =855
                    Height =315
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label548"
                    Caption ="Due On"
                    GroupTable =1
                    LayoutCachedLeft =15090
                    LayoutCachedTop =600
                    LayoutCachedWidth =15945
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =30
                    Top =975
                    Width =0
                    Height =315
                    Name ="EmptyCell663"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =975
                    LayoutCachedWidth =30
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =90
                    Top =975
                    Width =825
                    Height =315
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    ForeColor =10040879
                    Name ="Label20"
                    Caption ="♦BillID"
                    GroupTable =1
                    LayoutCachedLeft =90
                    LayoutCachedTop =975
                    LayoutCachedWidth =915
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =975
                    Top =975
                    Width =1008
                    Height =315
                    Name ="Label739"
                    Caption ="Sch"
                    GroupTable =1
                    LayoutCachedLeft =975
                    LayoutCachedTop =975
                    LayoutCachedWidth =1983
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2040
                    Top =975
                    Width =576
                    Height =315
                    Name ="Label740"
                    Caption ="Grp"
                    GroupTable =1
                    LayoutCachedLeft =2040
                    LayoutCachedTop =975
                    LayoutCachedWidth =2616
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2670
                    Top =975
                    Width =1872
                    Height =315
                    Name ="Label759"
                    Caption ="Contract"
                    GroupTable =1
                    LayoutCachedLeft =2670
                    LayoutCachedTop =975
                    LayoutCachedWidth =4542
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4605
                    Top =975
                    Width =1245
                    Height =315
                    Name ="Label634"
                    Caption ="Invoice Num"
                    GroupTable =1
                    LayoutCachedLeft =4605
                    LayoutCachedTop =975
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =7470
                    Top =975
                    Width =1275
                    Height =315
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =12
                    Name ="BillBankIDCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName FROM Banks WHERE (((Banks.BankType)=\"1"
                        "\")) ORDER BY Banks.BankShortName;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Link to Buyer"
                    ControlTipText ="This is the Assigned Bank for this Group"
                    GroupTable =1

                    LayoutCachedLeft =7470
                    LayoutCachedTop =975
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =1290
                    DisplayAsHyperlink =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =5910
                            Top =975
                            Width =1500
                            Height =315
                            Name ="Label716"
                            Caption ="Bil Type"
                            GroupTable =1
                            LayoutCachedLeft =5910
                            LayoutCachedTop =975
                            LayoutCachedWidth =7410
                            LayoutCachedHeight =1290
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8805
                    Top =975
                    Width =1155
                    Height =315
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =13
                    Name ="BilldtCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="The Date of this Bill"
                    GroupTable =1

                    LayoutCachedLeft =8805
                    LayoutCachedTop =975
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10020
                    Top =975
                    Width =1155
                    Height =315
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =14
                    Name ="BillDueDtCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="Due Date of this Bill"
                    GroupTable =1

                    LayoutCachedLeft =10020
                    LayoutCachedTop =975
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11235
                    Top =975
                    Width =150
                    Height =315
                    Name ="EmptyCell624"
                    GroupTable =1
                    LayoutCachedLeft =11235
                    LayoutCachedTop =975
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11445
                    Top =975
                    Width =1155
                    Height =315
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =15
                    Name ="BillFromCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="Start of the Billing Date Range"
                    GroupTable =1

                    LayoutCachedLeft =11445
                    LayoutCachedTop =975
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12660
                    Top =975
                    Width =1155
                    Height =315
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =16
                    Name ="BilltoCBO"
                    Format ="Short Date"
                    ControlTipText ="End of the Billing Date Range"
                    GroupTable =1

                    LayoutCachedLeft =12660
                    LayoutCachedTop =975
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13875
                    Top =975
                    Width =1155
                    Height =315
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =17
                    BorderColor =9211020
                    Name ="BillSentCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="Date bill was sent to Client"
                    GroupTable =1

                    LayoutCachedLeft =13875
                    LayoutCachedTop =975
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1872
                    Left =15090
                    Top =975
                    Width =855
                    Height =315
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =18
                    Name ="DueonRecCbo"
                    RowSourceType ="Value List"
                    RowSource ="True;False"
                    ColumnWidths ="1440"
                    Tag ="Medit"
                    ControlTipText ="Due on Receipt Y/N"
                    GroupTable =1

                    LayoutCachedLeft =15090
                    LayoutCachedTop =975
                    LayoutCachedWidth =15945
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4755
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =3
                    Name ="DeleteRecord"
                    Caption ="DeleteSelected"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete The Selected Bill / Record"
                    GroupTable =2
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000003255d654 ,
                        0x3255d6e73255d68d000000000000000000000000000000000000000000000000 ,
                        0x3255d6423255d67500000000000000000000000000000000000000003255d68d ,
                        0x3255d6ff3255d6ff3255d660000000000000000000000000000000003255d642 ,
                        0x3255d6f63255d64b00000000000000000000000000000000000000003255d61b ,
                        0x3255d6ea3255d6ff3255d6f93255d64500000000000000003255d6423255d6f6 ,
                        0x3255d6b100000000000000000000000000000000000000000000000000000000 ,
                        0x3255d63f3255d6f63255d6ff3255d6f63255d6423255d6423255d6f63255d6f0 ,
                        0x3255d61b0000000000000000000000000000000072727296727272f0727272a2 ,
                        0x000000003255d6363255d6e73255d6ff3255d6f63255d6f63255d6ff3255d64b ,
                        0x0000000000000000000000000000000000000000727272cc727272ff727272e1 ,
                        0x72727203000000003255d6213255d6f03255d6ff3255d6ff3255d6b700000000 ,
                        0x0000000000000000000000000000000000000000000000007272722a00000000 ,
                        0x000000003255d6213255d6c03255d6ff3255d6ff3255d6ff3255d6ff3255d6ab ,
                        0x3255d60c00000000000000000000000000000000727272187272723372727203 ,
                        0x3255d65a3255d6f33255d6ff3255d6ff3255d6873255d62a3255d6ae3255d6ff ,
                        0x3255d6de3255d63f0000000000000000000000007272727e727272ff72727212 ,
                        0x3255d6e43255d6ff3255d6ea3255d6570000000000000000000000003255d636 ,
                        0x3255d68d3255d62700000000000000000000000072727287727272ff72727269 ,
                        0x3255d6393255d6873255d60f0000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727293727272ff727272a5 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727299727272ff727272b1 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272a5727272ff727272ba ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272b1727272ff727272c3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272b7727272ff727272d2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272c3727272ff727272db ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4755
                    LayoutCachedTop =60
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =4290
                    Top =60
                    Width =330
                    Height =255
                    Name ="EmptyCell769"
                    GroupTable =2
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    LayoutCachedLeft =4290
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =4605
                    Top =300
                    Width =1245
                    Name ="EmptyCell770"
                    GroupTable =1
                    LayoutCachedLeft =4605
                    LayoutCachedTop =300
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =540
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =0
                    Height =255
                    FontSize =10
                    Name ="BillType"
                    ControlSource ="BillType"
                    GroupTable =1

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =30
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =90
                    Top =30
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =10040879
                    Name ="UnitID"
                    ControlSource ="BillID"
                    ControlTipText ="FA's Bill Id for the Bill / Invoice"
                    GroupTable =1

                    LayoutCachedLeft =90
                    LayoutCachedTop =30
                    LayoutCachedWidth =915
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =975
                    Top =30
                    Width =1008
                    Height =255
                    FontSize =10
                    TabIndex =2
                    BorderColor =14211288
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    ControlTipText ="Schedule for the Master Lease"
                    GroupTable =1

                    LayoutCachedLeft =975
                    LayoutCachedTop =30
                    LayoutCachedWidth =1983
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =30
                    Width =576
                    Height =255
                    FontSize =10
                    TabIndex =3
                    BorderColor =14211288
                    Name ="SGrpNO"
                    ControlSource ="SGrpNO"
                    ControlTipText ="Unit Group for this Schedule"
                    GroupTable =1

                    LayoutCachedLeft =2040
                    LayoutCachedTop =30
                    LayoutCachedWidth =2616
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2670
                    Top =30
                    Width =1872
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="BContract"
                    ControlSource ="BContract"
                    Tag ="Medit"
                    ControlTipText ="Contract Number for the Bank Agreeement"
                    GroupTable =1

                    LayoutCachedLeft =2670
                    LayoutCachedTop =30
                    LayoutCachedWidth =4542
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4605
                    Top =30
                    Width =1245
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="BillNo"
                    ControlSource ="BillNo"
                    ControlTipText ="Invoice Number Often the Bank sets this Value"
                    GroupTable =1

                    LayoutCachedLeft =4605
                    LayoutCachedTop =30
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =5910
                    Top =30
                    Width =1500
                    Height =255
                    FontSize =10
                    TabIndex =6
                    Name ="BillTypeID"
                    ControlSource ="BillTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_BillType.BillTypeID, Lookup_BillType.BillType FROM Lookup_BillType"
                        ";"
                    ColumnWidths ="0;1872"
                    StatusBarText ="Link to Buyer"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="This is the Buyer, Double Click to view"
                    GroupTable =1

                    LayoutCachedLeft =5910
                    LayoutCachedTop =30
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =285
                    DisplayAsHyperlink =1
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =7470
                    Top =30
                    Width =1275
                    Height =255
                    FontSize =10
                    TabIndex =7
                    Name ="BillBank"
                    ControlSource ="BillBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankType FROM Banks WHERE (((Ban"
                        "ks.BankType)=\"1\")) ORDER BY Banks.BankShortName;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Link to Buyer"
                    ControlTipText ="This is the Assigned Bank for this Group"
                    GroupTable =1

                    LayoutCachedLeft =7470
                    LayoutCachedTop =30
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =285
                    DisplayAsHyperlink =1
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8805
                    Top =30
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =8
                    BackColor =15070975
                    Name ="Billdt"
                    ControlSource ="Billdt"
                    Format ="Short Date"
                    ControlTipText ="The Date of this Bill"
                    GroupTable =1

                    LayoutCachedLeft =8805
                    LayoutCachedTop =30
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10020
                    Top =30
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =9
                    BackColor =15070975
                    Name ="BillDueDt"
                    ControlSource ="BillDueDt"
                    Format ="Short Date"
                    ValidationRule =">=0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Due Date of this Bill"
                    GroupTable =1

                    LayoutCachedLeft =10020
                    LayoutCachedTop =30
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11235
                    Top =30
                    Width =75
                    Height =255
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =10
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="BillSchRef"
                    GroupTable =1

                    LayoutCachedLeft =11235
                    LayoutCachedTop =30
                    LayoutCachedWidth =11310
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11370
                    Top =30
                    Width =15
                    Height =255
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =11
                    Name ="GroupId"
                    ControlSource ="GroupId"
                    GroupTable =1

                    LayoutCachedLeft =11370
                    LayoutCachedTop =30
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11445
                    Top =30
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =12
                    Name ="Billfrom"
                    ControlSource ="Billfrom"
                    Format ="Short Date"
                    ControlTipText ="Start of the Billing Date Range"
                    GroupTable =1

                    LayoutCachedLeft =11445
                    LayoutCachedTop =30
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12660
                    Top =30
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =13
                    Name ="Billto"
                    ControlSource ="Billto"
                    Format ="Short Date"
                    DefaultValue ="0"
                    ControlTipText ="End of the Billing Date Range"
                    GroupTable =1

                    LayoutCachedLeft =12660
                    LayoutCachedTop =30
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13875
                    Top =30
                    Width =1155
                    Height =255
                    FontSize =10
                    TabIndex =14
                    BackColor =15592953
                    BorderColor =9211020
                    Name ="BillSent"
                    ControlSource ="BillSent"
                    Format ="Short Date"
                    ControlTipText ="Date bill was sent to Client"
                    GroupTable =1

                    LayoutCachedLeft =13875
                    LayoutCachedTop =30
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1872
                    Left =15090
                    Top =30
                    Width =855
                    Height =255
                    FontSize =10
                    TabIndex =15
                    BackColor =15592953
                    Name ="DueonRec"
                    ControlSource ="DueonRec"
                    RowSourceType ="Value List"
                    RowSource ="True;False"
                    ColumnWidths ="1440"
                    StatusBarText ="Unit Status"
                    ControlTipText ="Due on Receipt Y/N"
                    Format ="True/False"
                    GroupTable =1

                    LayoutCachedLeft =15090
                    LayoutCachedTop =30
                    LayoutCachedWidth =15945
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
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
                    Width =15840
                    Height =59
                    FontSize =10
                    TabIndex =16
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
                    LayoutCachedWidth =15840
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
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BatchBillingMassModSub.cls"
