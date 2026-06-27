Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7470
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =10755
    Top =5460
    Right =27705
    Bottom =10320
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x73f1a382d14fe440
    End
    RecordSource ="SchOrderchq"
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
        Begin FormHeader
            Height =690
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =105
                    Width =6720
                    Height =585
                    Name ="Text5"
                    ControlSource ="=\"CHANGE SCHEDULE SORT ORDER FOR \" & [CLIENTSHNM] & \" \" & [MLNO]"

                    LayoutCachedTop =105
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =690
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6765
                    Top =210
                    Width =705
                    Height =390
                    TabIndex =1
                    Name ="SOFClBtn"
                    Caption ="Close "
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"SOFClBtn\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Stat"
                                "ements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6765
                    LayoutCachedTop =210
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =600
                End
            End
        End
        Begin Section
            Height =345
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1245
                    Width =930
                    Height =315
                    Name ="SCHNo1"
                    ControlSource ="SCHNo1"
                    Format ="Fixed"
                    StatusBarText ="added 8/15/2013 schedule number for sequencing"

                    LayoutCachedLeft =1245
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1050
                            Height =315
                            Name ="Label2"
                            Caption ="ORDER NO"
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3405
                    Width =3300
                    Height =315
                    TabIndex =1
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    StatusBarText ="schedule number starts with 1"

                    LayoutCachedLeft =3405
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2535
                            Width =765
                            Height =315
                            Name ="Label3"
                            Caption ="Sch No:"
                            LayoutCachedLeft =2535
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =315
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =45
            Name ="FormFooter"
        End
    End
End
