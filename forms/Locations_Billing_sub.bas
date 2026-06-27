Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4215
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =10440
    Top =6225
    Right =14265
    Bottom =7905
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x5516c0c61c07e540
    End
    RecordSource ="SELECT Contacts.contactID, Contacts.clientRef, Contacts.contactClassification, C"
        "ontacts.contactPrefix, Contacts.contactFirst, Contacts.contactMiddle, Contacts.c"
        "ontactLast, Contacts.contactSuffix, Contacts.contactTitle, Contacts.contactPrima"
        "ryPhone, Contacts.contactPhoneExt, Contacts.contactPhone2, Contacts.contactFax, "
        "Contacts.contactCell, Contacts.contactEmail, Contacts.contactFullName, Contacts."
        "contactAddress1, Contacts.contactAddress2, Contacts.contactCity, Contacts.LState"
        ", Contacts.contactZip, Contacts.LocationName FROM Contacts WHERE (((Contacts.con"
        "tactClassification)=\"Location\"));"
    Caption ="Locations"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =255
            BackColor =16249583
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Width =1140
                    ColumnWidth =570
                    ForeColor =1643706
                    Name ="contactID"
                    ControlSource ="contactID"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1335
                    Height =255
                    TabIndex =1
                    ForeColor =1643706
                    Name ="ClientREF"
                    ControlSource ="ClientREF"
                    StatusBarText ="link to clients table"

                    LayoutCachedLeft =1335
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =255
                End
            End
        End
        Begin Section
            Height =2145
            BackColor =16249583
            Name ="Detail"
            AlternateBackColor =15523798
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =975
                    Top =15
                    Width =3000
                    Height =255
                    ColumnWidth =2745
                    Name ="Street"
                    ControlSource ="contactAddress1"

                    LayoutCachedLeft =975
                    LayoutCachedTop =15
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Width =780
                            Height =240
                            Name ="Label25"
                            Caption ="Address:"
                            LayoutCachedLeft =165
                            LayoutCachedWidth =945
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =975
                    Top =600
                    Width =1260
                    Height =255
                    ColumnWidth =1650
                    TabIndex =1
                    Name ="City"
                    ControlSource ="contactCity"

                    LayoutCachedLeft =975
                    LayoutCachedTop =600
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =855
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2895
                    Top =600
                    Width =1080
                    Height =255
                    ColumnWidth =1335
                    TabIndex =3
                    Name ="Zip"
                    ControlSource ="contactZip"

                    LayoutCachedLeft =2895
                    LayoutCachedTop =600
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =855
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2265
                    Top =600
                    Width =600
                    Height =255
                    ColumnWidth =1125
                    TabIndex =2
                    Name ="State"
                    ControlSource ="LState"
                    StatusBarText ="lookup in lookup_state table"

                    LayoutCachedLeft =2265
                    LayoutCachedTop =600
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =855
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1005
                    Top =975
                    Width =2985
                    Height =255
                    TabIndex =4
                    Name ="contactFullName"
                    ControlSource ="contactFullName"

                    LayoutCachedLeft =1005
                    LayoutCachedTop =975
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =975
                            Width =780
                            Height =240
                            Name ="Label20"
                            Caption ="Contact:"
                            LayoutCachedLeft =165
                            LayoutCachedTop =975
                            LayoutCachedWidth =945
                            LayoutCachedHeight =1215
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1005
                    Top =1275
                    Width =2985
                    Height =255
                    TabIndex =5
                    Name ="Email"
                    ControlSource ="contactEmail"

                    LayoutCachedLeft =1005
                    LayoutCachedTop =1275
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1530
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =1275
                            Width =780
                            Height =240
                            Name ="Label21"
                            Caption ="Email:"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1275
                            LayoutCachedWidth =945
                            LayoutCachedHeight =1515
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1005
                    Top =1575
                    Width =1155
                    Height =255
                    TabIndex =6
                    Name ="Phone"
                    ControlSource ="contactPrimaryPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =1005
                    LayoutCachedTop =1575
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =1575
                            Width =780
                            Height =240
                            Name ="Label22"
                            Caption ="Phone:"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1575
                            LayoutCachedWidth =945
                            LayoutCachedHeight =1815
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1005
                    Top =1890
                    Width =1155
                    Height =255
                    TabIndex =7
                    Name ="Fax"
                    ControlSource ="contactFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =1005
                    LayoutCachedTop =1890
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =2145
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =1890
                            Width =780
                            Height =240
                            Name ="Fax Label"
                            Caption ="Fax:"
                            EventProcPrefix ="Fax_Label"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1890
                            LayoutCachedWidth =945
                            LayoutCachedHeight =2130
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =975
                    Top =300
                    Width =2999
                    Height =255
                    TabIndex =8
                    Name ="contactAddress2"
                    ControlSource ="contactAddress2"

                    LayoutCachedLeft =975
                    LayoutCachedTop =300
                    LayoutCachedWidth =3974
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2205
                    Top =1575
                    Width =945
                    Height =255
                    TabIndex =9
                    Name ="contactPhoneExt"
                    ControlSource ="contactPhoneExt"

                    LayoutCachedLeft =2205
                    LayoutCachedTop =1575
                    LayoutCachedWidth =3150
                    LayoutCachedHeight =1830
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =5
                    Left =945
                    Top =900
                    Width =3090
                    Name ="Line28"
                    LayoutCachedLeft =945
                    LayoutCachedTop =900
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =900
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15523798
            Name ="FormFooter"
        End
    End
End
