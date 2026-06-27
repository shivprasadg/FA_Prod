Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12870
    DatasheetFontHeight =10
    ItemSuffix =78
    Left =3405
    Top =5700
    Right =17175
    Bottom =12915
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x4174cc2af383e340
    End
    RecordSource ="Pricing"
    Caption ="Lease Pricing"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    SplitFormOrientation =1
    NavigationCaption ="Lease Pricing"
    SplitFormOrientation =1
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Visible = NotDefault
            Height =435
            Name ="FormHeader"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =960
                    Top =15
                    Width =780
                    Height =420
                    ForeColor =255
                    Name ="groupRef"
                    ControlSource ="groupRef"
                    StatusBarText ="link to Groups table"

                    LayoutCachedLeft =960
                    LayoutCachedTop =15
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =435
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =810
                            Height =435
                            Name ="Label65"
                            Caption ="groupRef:"
                            LayoutCachedWidth =810
                            LayoutCachedHeight =435
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6525
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6165
                    Top =105
                    Width =855
                    Height =255
                    BorderColor =13017476
                    Name ="pricingEconSpread"
                    ControlSource ="pricingEconSpread"

                    LayoutCachedLeft =6165
                    LayoutCachedTop =105
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5160
                            Top =105
                            Width =915
                            Height =240
                            Name ="Label72"
                            Caption ="Spread:"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =105
                            LayoutCachedWidth =6075
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8085
                    Top =105
                    Width =915
                    Height =255
                    TabIndex =1
                    BorderColor =13017476
                    Name ="pricingEconYield"
                    ControlSource ="pricingEconYield"

                    LayoutCachedLeft =8085
                    LayoutCachedTop =105
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7560
                            Top =105
                            Width =435
                            Height =240
                            Name ="Label73"
                            Caption ="Yield:"
                            LayoutCachedLeft =7560
                            LayoutCachedTop =105
                            LayoutCachedWidth =7995
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3750
                    Top =60
                    Width =870
                    Height =390
                    FontSize =9
                    TabIndex =2
                    Name ="RefreshBtn"
                    Caption ="Refresh"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="RunCommand"
                            Argument ="18"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"RefreshBtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"RunMenuCommand\"><A"
                                "rgument Name=\"Command\">Refresh<"
                        End
                        Begin
                            Comment ="_AXL:/Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3750
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =450
                End
                Begin Subform
                    OverlapFlags =87
                    Left =705
                    Top =570
                    Width =10920
                    Height =5715
                    TabIndex =3
                    Name ="QuoteFinal_sub"
                    SourceObject ="Form.QuoteFinal_sub"
                    LinkChildFields ="groupRef"
                    LinkMasterFields ="groupRef"

                    LayoutCachedLeft =705
                    LayoutCachedTop =570
                    LayoutCachedWidth =11625
                    LayoutCachedHeight =6285
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =705
                            Top =330
                            Width =1200
                            Height =240
                            Name ="QuoteFinal_sub Label"
                            Caption ="QuoteFinal_sub"
                            EventProcPrefix ="QuoteFinal_sub_Label"
                            LayoutCachedLeft =705
                            LayoutCachedTop =330
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =570
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1470
                    Width =2040
                    Height =390
                    TabIndex =4
                    Name ="AddQuoteBtn"
                    Caption ="add Quote"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="QuoteFinal_sub"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![QuoteFinal_sub]![groupRef]"
                            Argument ="[lease_Pricing_sub]![groupref]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddQuoteBtn\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argume"
                                "nt Name=\"FormName\">QuoteFinal_s"
                        End
                        Begin
                            Comment ="_AXL:ub</Argument><Argument Name=\"DataMode\">Add</Argument></Action><Action Nam"
                                "e=\"SetValue\"><Argument Name=\"Item\">[Forms]![QuoteFinal_sub]![groupRef]</Argu"
                                "ment><Argument Name=\"Expression\">[lease_Pricing_sub]![groupref]</Argument></Ac"
                                "tion></Statements></U"
                        End
                        Begin
                            Comment ="_AXL:serInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =1470
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =390
                    Overlaps =1
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
' See "Lease_Pricing_sub.cls"
