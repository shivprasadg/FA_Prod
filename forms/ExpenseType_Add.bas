Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5400
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =3405
    Top =2550
    Right =9345
    Bottom =7680
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xd3910eb22fa3e340
    End
    RecordSource ="Lookup_ExpenseTypes"
    Caption ="Add or Edit an Expense Type"
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
            Height =390
            Name ="FormHeader"
            AutoHeight =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =4755
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Add/Edit an Expense Type"
                    FontName ="Tahoma"
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =390
                End
            End
        End
        Begin Section
            Height =1380
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1575
                    Top =240
                    Width =3405
                    Height =315
                    FontSize =10
                    Name ="ExpenseType"
                    ControlSource ="ExpenseType"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1575
                    LayoutCachedTop =240
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =240
                            Width =1410
                            Height =315
                            FontSize =10
                            Name ="Label0"
                            Caption ="Expense Type:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =60
                            LayoutCachedTop =240
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Top =1020
                    Width =1200
                    Height =330
                    FontSize =9
                    TabIndex =1
                    ForeColor =255
                    Name ="CloseFormBtn"
                    Caption ="Save && Close"
                    FontName ="Tahoma"
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
                                "nterfaceMacro For=\"CloseFormBtn\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></"
                                "Statements></UserInterfaceMacro"
                        End
                        Begin
                            Comment ="_AXL:>"
                        End
                    End

                    LayoutCachedTop =1020
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =1350
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1320
                    Top =1020
                    Width =1770
                    Height =330
                    FontSize =9
                    TabIndex =2
                    ForeColor =5026082
                    Name ="EditExpTypeBtn"
                    Caption ="Edit Expense Types"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="ShowAllRecords"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"EditExpTypeBtn\" xmlns=\"http://schemas.microsoft.com/office"
                                "/accessservices/2009/11/application\"><Statements><Action Name=\"ShowAllRecords\""
                                "/></Statements></UserInterface"
                        End
                        Begin
                            Comment ="_AXL:Macro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1320
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1350
                End
            End
        End
        Begin FormFooter
            Height =120
            Name ="FormFooter"
        End
    End
End
