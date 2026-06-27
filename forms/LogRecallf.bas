Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6165
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =7275
    Top =6495
    Right =19470
    Bottom =18900
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x382b0476e738e440
    End
    RecordSource ="SELECT Recalls.RCID, Recalls.RCUnitID, Recalls.recallref FROM Recalls;"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OnLoadEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="MoveSize"
            Argument ="4320"
            Argument ="4320"
            Argument ="7200"
            Argument ="4320"
        End
    End
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
        Begin Section
            Height =915
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =1125
                    Top =480
                    Width =1185
                    Height =210
                    FontSize =9
                    ForeColor =5026082
                    Name ="RCUnitID"
                    ControlSource ="RCUnitID"
                    StatusBarText ="Xref to Unit being recalled"

                    LayoutCachedLeft =1125
                    LayoutCachedTop =480
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =15
                            Top =465
                            Width =960
                            Height =225
                            FontSize =9
                            ForeColor =5026082
                            Name ="Label0"
                            Caption ="RCUnitID:"
                            LayoutCachedLeft =15
                            LayoutCachedTop =465
                            LayoutCachedWidth =975
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =4530
                    Top =480
                    Width =1185
                    Height =210
                    FontSize =9
                    TabIndex =1
                    ForeColor =5026082
                    Name ="recallref"
                    ControlSource ="recallref"

                    LayoutCachedLeft =4530
                    LayoutCachedTop =480
                    LayoutCachedWidth =5715
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =3450
                            Top =465
                            Width =915
                            Height =225
                            FontSize =9
                            ForeColor =5026082
                            Name ="Label1"
                            Caption ="recallref:"
                            LayoutCachedLeft =3450
                            LayoutCachedTop =465
                            LayoutCachedWidth =4365
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Width =2640
                    Height =345
                    FontSize =14
                    FontWeight =700
                    Name ="Label9"
                    Caption ="ADD RECALL TO UNIT"
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =345
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4950
                    Width =885
                    Height =390
                    TabIndex =2
                    Name ="Command12"
                    Caption ="Close"
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
                                "nterfaceMacro For=\"Command12\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4950
                    LayoutCachedWidth =5835
                    LayoutCachedHeight =390
                End
            End
        End
    End
End
