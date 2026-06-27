Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4485
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =9420
    Top =6120
    Right =14895
    Bottom =8550
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xdf35237db2fde340
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
        Begin Section
            Height =1020
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Width =1800
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Unlock Group"
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =390
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =660
                    Width =1125
                    Height =315
                    ForeColor =16777215
                    Name ="PWord"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =660
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =975
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =435
                            Top =660
                            Width =1020
                            Height =315
                            Name ="Label2"
                            Caption ="Password:"
                            LayoutCachedLeft =435
                            LayoutCachedTop =660
                            LayoutCachedWidth =1455
                            LayoutCachedHeight =975
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =660
                    Width =1191
                    Height =315
                    TabIndex =1
                    Name ="unlock"
                    Caption ="Unlock"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =660
                    LayoutCachedWidth =4191
                    LayoutCachedHeight =975
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =105
                    Width =1191
                    Height =315
                    TabIndex =2
                    Name ="CLGrpULBtn"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =105
                    LayoutCachedWidth =4191
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "Unlock Grp.cls"
