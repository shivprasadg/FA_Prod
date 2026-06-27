Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =4365
    Top =5760
    Right =11955
    Bottom =10590
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x19e0e4939561e440
    End
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
            FontName ="Calibri"
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
        Begin FormHeader
            Height =495
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =165
                    Top =75
                    Width =6390
                    Height =405
                    FontSize =14
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label8"
                    Caption ="Check out ALL Titles on Schedule To Trust / Custodian"
                    LayoutCachedLeft =165
                    LayoutCachedTop =75
                    LayoutCachedWidth =6555
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            Height =2385
            Name ="Detail"
            Begin
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =2625
                    Top =195
                    Width =3015
                    Height =315
                    Name ="SelClient"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm, Clients.[CLIENT GROUP] FROM Clients WHERE (((Clients."
                        "ClientShNm)<>\"demo a\" And (Clients.ClientShNm)<>\"demo b\" And (Clients.Client"
                        "ShNm)<>\"demo c\" And (Clients.ClientShNm)<>\"demo d\" And (Clients.ClientShNm)<"
                        ">\"demo e\" And (Clients.ClientShNm)<>\"demo f\" And (Clients.ClientShNm)<>\"dem"
                        "o g\" And (Clients.ClientShNm)<>\"demo h\") AND ((Clients.clientType)=\"client\""
                        ") AND ((Clients.[CLIENT GROUP])<>\"client demo\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="1440"
                    Format =">"
                    OnDirty ="Refresh"
                    GroupTable =1
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"SelClient\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"Requery\"/></Stateme"
                                "nts></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =2625
                    LayoutCachedTop =195
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =510
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =510
                            Top =195
                            Width =2055
                            Height =315
                            Name ="ClientShNm_Label"
                            Caption ="1. Select Customer"
                            GroupTable =1
                            LayoutCachedLeft =510
                            LayoutCachedTop =195
                            LayoutCachedWidth =2565
                            LayoutCachedHeight =510
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =2625
                    Top =570
                    Width =3015
                    Height =315
                    TabIndex =1
                    Name ="SelLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLNo FROM Clients INNER JOIN MstrLease ON Clients.clientID = Ms"
                        "trLease.Client WHERE (((Clients.ClientShNm)=[forms]![TitletoTrustq]![selclient])"
                        " AND ((MstrLease.MLOrig)=\"fa\") AND ((Clients.[client group]) Not Like \"*demo*"
                        "\")) ORDER BY MstrLease.MLNo;"
                    ColumnWidths ="1440"
                    OnClick ="Refresh"
                    OnChange ="Refresh"
                    Format =">"
                    GroupTable =1

                    LayoutCachedLeft =2625
                    LayoutCachedTop =570
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =885
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =510
                            Top =570
                            Width =2055
                            Height =315
                            Name ="Label4"
                            Caption ="2. Select Lease"
                            GroupTable =1
                            LayoutCachedLeft =510
                            LayoutCachedTop =570
                            LayoutCachedWidth =2565
                            LayoutCachedHeight =885
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =4464
                    Left =2625
                    Top =945
                    Width =3015
                    Height =315
                    TabIndex =2
                    ForeColor =-2147483617
                    Name ="SelSch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo, IIf([schfsl]=True,\"yes\",Null) AS FSL, M"
                        "in(SchGrp.BLCD\015\012\015\012\015\012\015\012) AS MinOfBLCD\015\012\015\012\015"
                        "\012\015\012 FROM (MstrLease INNER JOIN Schedule ON MstrLease.MLID = Schedule.Sc"
                        "hMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID GROUP BY Schedule"
                        ".SchID, Schedule.SchNo, IIf([schfsl]=True,\"yes\",Null), MstrLease.MLNo HAVING ("
                        "((MstrLease.MLNo)=[forms]![TitletoTrustq]![selLease])) ORDER BY Schedule.SchID D"
                        "ESC , Schedule.SchNo DESC;"
                    ColumnWidths ="0;2160;864;1440"
                    OnClick ="Refresh"
                    OnChange ="Refresh"
                    Format ="General Number"
                    GroupTable =1

                    LayoutCachedLeft =2625
                    LayoutCachedTop =945
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1260
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =510
                            Top =945
                            Width =2055
                            Height =315
                            Name ="Label6"
                            Caption ="3. Select Schedule"
                            GroupTable =1
                            LayoutCachedLeft =510
                            LayoutCachedTop =945
                            LayoutCachedWidth =2565
                            LayoutCachedHeight =1260
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2625
                    Top =1320
                    Width =3015
                    Height =315
                    TabIndex =3
                    Name ="SentDate"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =2625
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1635
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =510
                            Top =1320
                            Width =2055
                            Height =315
                            Name ="Label7"
                            Caption ="4. Enter date to Trust:"
                            GroupTable =1
                            LayoutCachedLeft =510
                            LayoutCachedTop =1320
                            LayoutCachedWidth =2565
                            LayoutCachedHeight =1635
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =495
                    Top =1815
                    Width =5160
                    Height =390
                    TabIndex =4
                    Name ="ckoutschbtn"
                    Caption ="Check out Titles to Trust / Custodian"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =495
                    LayoutCachedTop =1815
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =2205
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5955
                    Top =240
                    Width =450
                    Height =315
                    TabIndex =5
                    Name ="out"
                    ControlSource ="2"
                    Format ="General Number"

                    LayoutCachedLeft =5955
                    LayoutCachedTop =240
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5955
                    Top =675
                    Width =450
                    Height =315
                    TabIndex =6
                    Name ="why"
                    ControlSource ="4"
                    Format ="General Number"

                    LayoutCachedLeft =5955
                    LayoutCachedTop =675
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =990
                End
            End
        End
        Begin FormFooter
            Height =210
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "TitletoTrustq.cls"
