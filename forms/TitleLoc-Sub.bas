Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5220
    DatasheetFontHeight =11
    ItemSuffix =7
    Left =4350
    Top =8145
    Right =9675
    Bottom =9630
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x036e2bca1c07e540
    End
    RecordSource ="SELECT Contacts.contactID, Contacts.contactAddress1, Contacts.contactAddress2, C"
        "ontacts.contactCity, Contacts.contactcounty, Contacts.LState, Contacts.contactZi"
        "p FROM Contacts WHERE (((Contacts.contactClassification)=\"LOCATIONS\"));"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Section
            Height =975
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Width =60
                    Height =315
                    ColumnWidth =1290
                    Name ="contactID"
                    ControlSource ="contactID"

                    LayoutCachedWidth =60
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Width =5220
                    Height =315
                    TabIndex =1
                    Name ="contactAddress1"
                    ControlSource ="contactAddress1"

                    LayoutCachedWidth =5220
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =330
                    Width =5220
                    Height =315
                    TabIndex =2
                    Name ="contactAddress2"
                    ControlSource ="contactAddress2"

                    LayoutCachedTop =330
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =645
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Top =660
                    Width =2220
                    Height =315
                    TabIndex =3
                    Name ="contactCity"
                    ControlSource ="contactCity"

                    LayoutCachedTop =660
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =2220
                    Top =660
                    Height =315
                    TabIndex =4
                    Name ="contactcounty"
                    ControlSource ="contactcounty"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =660
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4140
                    Top =660
                    Width =1080
                    Height =315
                    TabIndex =6
                    Name ="contactZip"
                    ControlSource ="contactZip"
                    InputMask ="00000\\-9999;;_"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =660
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3660
                    Top =660
                    Width =480
                    Height =315
                    TabIndex =5
                    Name ="LState"
                    ControlSource ="LState"
                    StatusBarText ="lookup in lookup_state table"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =660
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =975
                End
            End
        End
    End
End
