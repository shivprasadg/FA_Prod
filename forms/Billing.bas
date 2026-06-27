Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20880
    DatasheetFontHeight =11
    ItemSuffix =35
    Left =5670
    Top =6180
    Right =21285
    Bottom =17175
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc2bf92bfec70e440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =900
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13260
                    Top =405
                    Width =1215
                    ColumnOrder =0
                    FontSize =9
                    ForeColor =5855577
                    Name ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    GroupTable =1

                    LayoutCachedLeft =13260
                    LayoutCachedTop =405
                    LayoutCachedWidth =14475
                    LayoutCachedHeight =645
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14535
                    Top =405
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =1
                    ForeColor =5855577
                    Name ="SchNo"
                    GroupTable =1

                    LayoutCachedLeft =14535
                    LayoutCachedTop =405
                    LayoutCachedWidth =15975
                    LayoutCachedHeight =645
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18045
                    Top =405
                    Width =630
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =5
                    ForeColor =5855577
                    Name ="BillID"
                    GroupTable =1

                    LayoutCachedLeft =18045
                    LayoutCachedTop =405
                    LayoutCachedWidth =18675
                    LayoutCachedHeight =645
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16755
                    Top =405
                    Width =630
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =3
                    ForeColor =5855577
                    Name ="SchID"
                    GroupTable =1

                    LayoutCachedLeft =16755
                    LayoutCachedTop =405
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =645
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    Left =13275
                    Top =180
                    Width =5460
                    Height =240
                    FontSize =9
                    Name ="Label8"
                    Caption ="MLNo.               |Schedule No.       |ClientID| SCHID |GrpID |Bill ID        "
                        "              "
                    LayoutCachedLeft =13275
                    LayoutCachedTop =180
                    LayoutCachedWidth =18735
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17445
                    Top =405
                    Width =540
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =4
                    ForeColor =5855577
                    Name ="GroupId"
                    GroupTable =1

                    LayoutCachedLeft =17445
                    LayoutCachedTop =405
                    LayoutCachedWidth =17985
                    LayoutCachedHeight =645
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16035
                    Top =405
                    Width =660
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =2
                    ForeColor =5855577
                    Name ="clientId"
                    GroupTable =1

                    LayoutCachedLeft =16035
                    LayoutCachedTop =405
                    LayoutCachedWidth =16695
                    LayoutCachedHeight =645
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =720
                    Width =20880
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =720
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =900
                    TabIndex =10
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureAlignment =0
                    PictureType =2
                    Width =661
                    Height =780
                    Name ="Image1161"
                    Picture ="ATLAASGlobeLogo"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =661
                    LayoutCachedHeight =780
                    TabIndex =9
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =127
                    TextAlign =1
                    Left =810
                    Width =6660
                    Height =720
                    FontSize =24
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Billing - Client Invoicing"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =810
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =720
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =1050
                    Top =525
                    Width =840
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label227"
                    Caption ="Manager"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1050
                    LayoutCachedTop =525
                    LayoutCachedWidth =1890
                    LayoutCachedHeight =765
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =19410
                    Top =210
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19410
                    LayoutCachedTop =210
                    LayoutCachedWidth =19965
                    LayoutCachedHeight =438
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =19977
                    Top =210
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =7
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =19977
                    LayoutCachedTop =210
                    LayoutCachedWidth =20532
                    LayoutCachedHeight =438
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8160
                    Top =240
                    Width =3240
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="RunAutoExtend_"
                    Caption ="Auto Extend Expiring MTM Leases"
                    OnClick ="[Event Procedure]"
                    LeftPadding =105
                    TopPadding =90
                    RightPadding =120
                    BottomPadding =90
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727287727272f3727272f3 ,
                        0x72727287000000000000000000000000727272b7727272b70000000000000000 ,
                        0x0000000000000000000000000000000000000000727272fc727272ff727272ff ,
                        0x727272f9000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000072727203727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ba727272ff727272ff ,
                        0x7272729c000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x0000000000000000727272067272728d727272ff727272ff7272728d72727206 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x000000000000000072727287727272ff727272ff727272ff727272ff7272728a ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x0000000000000000727272ea727272ff727272bd727272bd727272ff727272ea ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff00000000727272ea727272ff0000000000000000727272ff727272ea ,
                        0x000000000000000000000000000000000000000072727230727272ff727272ff ,
                        0x727272300000000072727287727272ff0000000000000000727272ff72727287 ,
                        0x00000000000000000000000000000000000000007272720f727272ff727272ff ,
                        0x7272720f00000000727272067272728400000000000000007272728472727206 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8160
                    LayoutCachedTop =240
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =525
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =0
                    QuickStyle =7
                    QuickStyleMask =-193
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =10110
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =60
                    Top =300
                    Width =3120
                    Height =315
                    FontSize =10
                    BackColor =15858167
                    Name ="SelClient"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.clientID, Clients.ClientShNm, Clients.clientType FROM Clients WHE"
                        "RE (((Clients.clientType) Like \"*Client*\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="0;2880"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Format =">"
                    OnDirty ="Refresh"
                    AllowValueListEdits =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =615
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =45
                            Width =1800
                            Height =315
                            FontSize =10
                            Name ="ClientShNm_Label"
                            Caption ="1. Select Customer"
                            FontName ="Segoe UI"
                            LayoutCachedTop =45
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =360
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =60
                    Top =2940
                    Width =3120
                    Height =3360
                    FontSize =9
                    TabIndex =2
                    BoundColumn =1
                    ForeColor =-2147483617
                    Name ="SelSch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo, DateDiff(\"m\",[BLCD],[AssignmentDate]) A"
                        "S Strp, Format(Min(SchGrp.BLCD),'mm/dd/yy') AS BLCD, SchGrp.AssignmentDate FROM "
                        "(MstrLease INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef) INNER JOIN"
                        " SchGrp ON Schedule.SchID = SchGrp.SGrpSchID GROUP BY Schedule.SchID, Schedule.S"
                        "chNo, MstrLease.MLID, SchGrp.AssignmentDate HAVING (((MstrLease.MLID)=181)) ORDE"
                        "R BY CDbl(NumInString(Left(SchNo,3))) DESC , Schedule.SchNo DESC;"
                    ColumnWidths ="0;1656;374;1008"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =6300
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =2670
                            Width =1800
                            Height =315
                            FontSize =10
                            Name ="Label6"
                            Caption ="3. Select Schedule"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2670
                            LayoutCachedWidth =1860
                            LayoutCachedHeight =2985
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =3270
                    Width =17595
                    Height =9900
                    TabIndex =3
                    Name ="BillingSubForm"
                    SourceObject ="Form.BillingSubForm"
                    LinkChildFields ="MLNo;SchNo"
                    LinkMasterFields ="MLNo;SchNo"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =3270
                    LayoutCachedWidth =20865
                    LayoutCachedHeight =9900
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =9030
                    Width =3120
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="UpdateBilling"
                    Caption ="Update Billing Address/Bank Nos"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =9030
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =9330
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =945
                    Width =3120
                    Height =1560
                    FontSize =9
                    TabIndex =1
                    Name ="SelLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, Client From MstrLease WHERE MLOrig = 'fa' and Client=4971 ORD"
                        "ER BY MstrLease.MLNo;"
                    ColumnWidths ="0;2880"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =945
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2505
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =675
                            Width =1800
                            Height =315
                            FontSize =10
                            Name ="Label4"
                            Caption ="2. Select Lease"
                            FontName ="Segoe UI"
                            LayoutCachedTop =675
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =990
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =60
                    Top =6660
                    Width =3120
                    Height =2310
                    FontSize =9
                    TabIndex =5
                    BoundColumn =1
                    BackColor =15858167
                    ForeColor =-2147483617
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchGrp.SGrpID AS GroupId, SchGrp.SGrpNO AS [Group], SchGrp.UnitsInGroup A"
                        "S Units, Min(Format([SchGrp].[BLCD],\"mm/dd/yy\")) AS BLCD, Schedule.SchID FROM "
                        "Schedule INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID GROUP BY SchGrp."
                        "SGrpID, SchGrp.SGrpNO, SchGrp.UnitsInGroup, Schedule.SchID HAVING (((Schedule.Sc"
                        "hID)=[Forms]![Billing]![schid])) ORDER BY Schedule.SchID DESC;"
                    ColumnWidths ="0;1634;475;1008"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =6660
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =8970
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =6390
                            Width =2205
                            Height =315
                            FontSize =10
                            Name ="Label10"
                            Caption ="Option: Select A Group"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =6390
                            LayoutCachedWidth =2265
                            LayoutCachedHeight =6705
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                        End
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
' See "Billing.cls"
