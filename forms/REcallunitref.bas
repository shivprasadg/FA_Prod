Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5535
    DatasheetFontHeight =11
    ItemSuffix =3
    Left =5340
    Top =7170
    Right =16260
    Bottom =10290
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xb44d7a0ed738e440
    End
    RecordSource ="SELECT Recalls.RCID, Recalls.RCUnitID, Recalls.recallref FROM Recalls;"
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
            Height =1095
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =135
                    Height =315
                    Name ="RCUnitID"
                    ControlSource ="RCUnitID"
                    StatusBarText ="Xref to Unit being recalled"
                    DefaultValue ="[Forms]![recall checkin]![seekvin]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =135
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =450
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =135
                            Width =960
                            Height =315
                            Name ="Label0"
                            Caption ="RCUnitID:"
                            LayoutCachedTop =135
                            LayoutCachedWidth =960
                            LayoutCachedHeight =450
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =540
                    Height =315
                    TabIndex =1
                    Name ="RCNo"
                    ControlSource ="recallref"
                    StatusBarText ="recallcampaign table dnu"
                    DefaultValue ="[Forms]![recall checkin]![recallno]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =540
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =855
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =540
                            Width =645
                            Height =315
                            Name ="Label1"
                            Caption ="RCNo:"
                            LayoutCachedTop =540
                            LayoutCachedWidth =645
                            LayoutCachedHeight =855
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3675
                    Top =525
                    Height =315
                    TabIndex =2
                    Name ="RCID"
                    ControlSource ="RCID"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =525
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =840
                End
            End
        End
    End
End
