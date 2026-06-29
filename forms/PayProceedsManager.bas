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
    Width =28800
    DatasheetFontHeight =11
    ItemSuffix =101
    Left =5115
    Top =2535
    Right =-31366
    Bottom =17340
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
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
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ModernChart
            OldBorderStyle =0
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackThemeColorIndex =1
            ChartTitle ="Chart Title"
            ChartTitleFontName ="Calibri"
            ThemeFontIndex =1
            ChartSubtitle ="Subtitle"
        End
        Begin FormHeader
            Height =735
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4440
                    Top =300
                    Width =9000
                    Height =360
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    Name ="SchNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =300
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4440
                            Top =120
                            Width =1800
                            Height =180
                            FontSize =8
                            Name ="Label13"
                            Caption ="You're Working with:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =120
                            LayoutCachedWidth =6240
                            LayoutCachedHeight =300
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16470
                    Width =480
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =16470
                    LayoutCachedWidth =16950
                    LayoutCachedHeight =210
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =15975
                    Top =15
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label49"
                    Caption ="Units"
                    LayoutCachedLeft =15975
                    LayoutCachedTop =15
                    LayoutCachedWidth =16425
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16455
                    Top =225
                    Width =480
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="Text63"
                    ControlSource ="=[Units_List].[ItemsSelected].[Count]"

                    LayoutCachedLeft =16455
                    LayoutCachedTop =225
                    LayoutCachedWidth =16935
                    LayoutCachedHeight =435
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =15960
                    Top =240
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label64"
                    Caption ="Units"
                    LayoutCachedLeft =15960
                    LayoutCachedTop =240
                    LayoutCachedWidth =16410
                    LayoutCachedHeight =435
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =540
                    Width =28800
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =28800
                    LayoutCachedHeight =735
                    TabIndex =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13765
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=[AuditManager].[Form]![UnitCount]"

                    LayoutCachedLeft =13765
                    LayoutCachedTop =240
                    LayoutCachedWidth =14454
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =13740
                            Top =60
                            Width =750
                            Height =540
                            FontSize =8
                            BackColor =1643706
                            BorderColor =1643706
                            Name ="Label77"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =60
                            LayoutCachedWidth =14490
                            LayoutCachedHeight =600
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Top =60
                    Width =4365
                    Height =570
                    FontSize =20
                    Name ="Label31"
                    Caption ="Synd Proceeds Manager"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =630
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =22515
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =4
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =22515
                    LayoutCachedTop =120
                    LayoutCachedWidth =23265
                    LayoutCachedHeight =348
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
                    OverlapFlags =215
                    Left =21960
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =21960
                    LayoutCachedTop =120
                    LayoutCachedWidth =22515
                    LayoutCachedHeight =348
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13860
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =225
                    Width =3060
                    Height =270
                    FontSize =9
                    BorderColor =967423
                    Name ="SearchClient"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =225
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =495
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =120
                            Width =3060
                            Height =240
                            FontSize =10
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label668"
                            Caption ="Type Search HERE:"
                            LayoutCachedLeft =120
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =240
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8010
                    Top =5760
                    Width =5475
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label58"
                    Caption ="┌──────  Search ───────────┬────────────┬──────┬────────────┐"
                    LayoutCachedLeft =8010
                    LayoutCachedTop =5760
                    LayoutCachedWidth =13485
                    LayoutCachedHeight =5985
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3240
                    Top =360
                    Width =4014
                    FontSize =9
                    TabIndex =1
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys WHERE MlOrig='FA' GROUP BY M"
                        "LID, MLNo, MLOrig, ClientID Having ClientID =4268 ORDER BY MLNo;"
                    ColumnWidths ="0;1872;576"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3240
                    LayoutCachedTop =360
                    LayoutCachedWidth =7254
                    LayoutCachedHeight =1800
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =3240
                    Top =2700
                    Width =4074
                    Height =10920
                    FontSize =9
                    TabIndex =11
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select SchID, SCH, BankShortName as Assignee, Units as Unt,uPayInterim as iUnts,"
                        "cIFParts as iPrts  From vw_PayProceeds_Sch_Units as P Inner Join Banks as B on P"
                        ".AssigneeID = B.bankID Where MLID=63  ORDER BY CDbl(NumInString(Left(Sch,3))) de"
                        "sc, Sch desc;"
                    ColumnWidths ="0;1008;1152;576;576;5760"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2700
                    LayoutCachedWidth =7314
                    LayoutCachedHeight =13620
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3240
                    Top =15
                    Width =1350
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label22"
                    Caption ="Leases:"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =15
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    Left =3240
                    Top =1860
                    Width =1350
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Schedules:"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =2160
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Left =11835
                    Top =5790
                    Width =691
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearLabel"
                    Caption ="◄Clear►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =11835
                    LayoutCachedTop =5790
                    LayoutCachedWidth =12526
                    LayoutCachedHeight =6015
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =15
                    Left =7500
                    Top =6360
                    Width =19380
                    Height =7500
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    BackColor =15527148
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    RowSource ="Select AssetID,UnitNum,VIN,ItemDesc,[Sch|Grp],Vendor,IFBank,VendorCost,uIFAmt,iF"
                        "Pts,PartCost,PartIFAmt,IFAmount,uIFVar ,cIFVar  FROM vw_PayProceeds_IfUnits as P"
                        "PIF WHERE SchID In (2398) ORDER BY IFBank DESC;"
                    ColumnWidths ="720;1008;2016;3312;1440;1152;864;1152;1152;432;1152;1152;1008;1008"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =7500
                    LayoutCachedTop =6360
                    LayoutCachedWidth =26880
                    LayoutCachedHeight =13860
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8232
                    Top =6012
                    Width =1008
                    FontSize =9
                    TabIndex =14
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchUnit"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8232
                    LayoutCachedTop =6012
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =6252
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9252
                    Top =6012
                    Width =2016
                    FontSize =9
                    TabIndex =15
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9252
                    LayoutCachedTop =6012
                    LayoutCachedWidth =11268
                    LayoutCachedHeight =6252
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7500
                    Top =6012
                    Width =720
                    FontSize =9
                    TabIndex =13
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchAsset"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7500
                    LayoutCachedTop =6012
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =6252
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11280
                    Top =6012
                    Width =288
                    FontSize =9
                    TabIndex =16
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11280
                    LayoutCachedTop =6012
                    LayoutCachedWidth =11568
                    LayoutCachedHeight =6252
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12744
                    Top =6012
                    Width =828
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchType"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12744
                    LayoutCachedTop =6012
                    LayoutCachedWidth =13572
                    LayoutCachedHeight =6252
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =120
                    Top =495
                    Width =3060
                    Height =13125
                    FontSize =9
                    TabIndex =2
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] Fr"
                        "om vw_SixKeys WHERE MLOrig='FA' GROUP BY clientID, ClientShortName, CompanyName "
                        "ORDER BY CompanyName;"
                    ColumnWidths ="0;2160;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =495
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =13620
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =14700
                    Top =2160
                    Width =3414
                    Height =3600
                    FontSize =9
                    TabIndex =6
                    Name ="DrawList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select DrawDownID, DrawNumber AS Num, DrawName, PPCertSent as CertSent From Draw"
                        "Downs WHERE DrawDownID IN  (Select DrawDownID     FROM vw_PayProceeds_Base AS PP"
                        "B     WHERE Schid In (2398)     Group BY DrawDownID);"
                    ColumnWidths ="0;432;1872;1008"
                    AllowValueListEdits =0

                    LayoutCachedLeft =14700
                    LayoutCachedTop =2160
                    LayoutCachedWidth =18114
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =14700
                            Top =1860
                            Width =2634
                            Height =300
                            FontSize =12
                            FontWeight =700
                            Name ="Label75"
                            Caption ="Draws:"
                            LayoutCachedLeft =14700
                            LayoutCachedTop =1860
                            LayoutCachedWidth =17334
                            LayoutCachedHeight =2160
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =2580
                    Top =15
                    Width =600
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="ClearClientSearch"
                    Caption ="▼Clear"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =2580
                    LayoutCachedTop =15
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11580
                    Top =6012
                    Width =1152
                    FontSize =9
                    TabIndex =17
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchMake"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11580
                    LayoutCachedTop =6012
                    LayoutCachedWidth =12732
                    LayoutCachedHeight =6252
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =3240
                    Top =2220
                    Width =1260
                    Height =420
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8355711
                    Name ="PickModeToggle"
                    DefaultValue ="0"
                    Caption ="▼Pick On"
                    FontName ="0ystem"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2220
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =2640
                    ForeTint =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15060409
                    BackTint =40.0
                    OldBorderStyle =1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =8
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =2220
                    Width =2754
                    Height =420
                    FontSize =12
                    TabIndex =4
                    Name ="GetSelectedUnits"
                    Caption ="▼Get Schedules ▼"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7314
                    LayoutCachedHeight =2640
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverColor =10319446
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =7440
                    Top =2160
                    Width =7194
                    Height =3600
                    FontSize =9
                    TabIndex =5
                    Name ="PayOutList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Sch ,iIF(Units> 0,'Unit','Part') As [Type] ,Sum(vUnits + vParts) as vCnt "
                        ",Sum(PayVendor) as ToVendor ,[InterimBank] as iFBank ,Sum((iUnits + iParts)) as "
                        "iCnt ,Sum(PPS.InterimAmt) as InterimAmt ,Sum((faUnits + faParts)) as faCnt ,Sum("
                        "[FAAmount]) as [FA Cash] From vw_PayProceeds_Summary as PPS WHERE Schid In (2398"
                        ") GROUP BY Sch, iIF(Units> 0,'Unit','Part'), [InterimBank] ORDER BY Sch, iIF(Uni"
                        "ts> 0,'Unit','Part') DESC,[InterimBank];"
                    ColumnWidths ="720;576;432;1152;1008;432;1152;432;1152"
                    AllowValueListEdits =0

                    LayoutCachedLeft =7440
                    LayoutCachedTop =2160
                    LayoutCachedWidth =14634
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7440
                            Top =1860
                            Width =2634
                            Height =300
                            FontSize =12
                            FontWeight =700
                            Name ="Label88"
                            Caption ="Pay Outs:"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1860
                            LayoutCachedWidth =10074
                            LayoutCachedHeight =2160
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23280
                    Top =6000
                    Height =315
                    FontSize =10
                    TabIndex =12
                    BackColor =16381933
                    Name ="UnitIF"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =23280
                    LayoutCachedTop =6000
                    LayoutCachedWidth =24720
                    LayoutCachedHeight =6315
                    ForeThemeColorIndex =0
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =18360
                    Top =3540
                    Width =2580
                    Height =288
                    FontSize =8
                    TabIndex =9
                    Name ="MakePPLetter"
                    Caption ="Make Pay Procceeds Letter"
                    OnClick ="[Event Procedure]"
                    GroupTable =4

                    LayoutCachedLeft =18360
                    LayoutCachedTop =3540
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =3828
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverShade =75.0
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =18360
                    Top =2880
                    Width =2580
                    Height =288
                    FontSize =8
                    TabIndex =8
                    Name ="MakePPCert"
                    Caption ="Make Pay Procceeds Cert"
                    OnClick ="[Event Procedure]"
                    GroupTable =4

                    LayoutCachedLeft =18360
                    LayoutCachedTop =2880
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =3168
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    PressedThemeColorIndex =4
                    PressedTint =40.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =18360
                    Top =4248
                    Width =2580
                    Height =240
                    FontSize =8
                    TabIndex =10
                    Name ="ExcelUnitPartReport"
                    Caption ="Drawn Unit/Part"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    GroupTable =4
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18360
                    LayoutCachedTop =4248
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =4488
                    PictureCaptionArrangement =5
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =2
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
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =18360
                    Top =2220
                    Width =2580
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    Name ="Editdraw"
                    Caption ="Edit Update  Draw"
                    OnClick ="[Event Procedure]"
                    GroupTable =4

                    LayoutCachedLeft =18360
                    LayoutCachedTop =2220
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =2505
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =1
                    HoverColor =48840
                    PressedColor =48840
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =18360
                    Top =3888
                    Width =2580
                    Height =288
                    Name ="EmptyCell98"
                    GroupTable =4
                    LayoutCachedLeft =18360
                    LayoutCachedTop =3888
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =4176
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =18360
                    Top =2568
                    Width =2580
                    Name ="EmptyCell99"
                    GroupTable =4
                    LayoutCachedLeft =18360
                    LayoutCachedTop =2568
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =2808
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =18360
                    Top =3228
                    Width =2580
                    Name ="EmptyCell100"
                    GroupTable =4
                    LayoutCachedLeft =18360
                    LayoutCachedTop =3228
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =3468
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
            End
        End
        Begin FormFooter
            Height =225
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18030
                    Width =705
                    Height =225
                    FontSize =9
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="SchID"

                    LayoutCachedLeft =18030
                    LayoutCachedWidth =18735
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =17460
                            Width =570
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label19"
                            Caption ="SchID:"
                            LayoutCachedLeft =17460
                            LayoutCachedWidth =18030
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19530
                    Width =705
                    Height =225
                    FontSize =9
                    TabIndex =1
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="GroupID"

                    LayoutCachedLeft =19530
                    LayoutCachedWidth =20235
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =18735
                            Width =795
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label81"
                            Caption ="GroupID:"
                            LayoutCachedLeft =18735
                            LayoutCachedWidth =19530
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21015
                    Width =705
                    Height =225
                    FontSize =9
                    TabIndex =2
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="UnitId"

                    LayoutCachedLeft =21015
                    LayoutCachedWidth =21720
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =20220
                            Width =795
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label85"
                            Caption ="UnitID:"
                            LayoutCachedLeft =20220
                            LayoutCachedWidth =21015
                            LayoutCachedHeight =225
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PayProceedsManager.cls"
