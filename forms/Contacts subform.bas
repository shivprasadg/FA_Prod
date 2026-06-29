Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8892
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =5310
    Top =6720
    Right =14490
    Bottom =10530
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x09a624c71c07e540
    End
    RecordSource ="SELECT [Contacts].[contactID], [Contacts].[contactAddress2], [Contacts].[contact"
        "Address1], [Contacts].[contactCity], [Contacts].[contactcounty], [Contacts].[LSt"
        "ate], [Contacts].[contactZip] FROM [Contacts]"
    Caption ="Contacts subform"
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
            TextFontCharSet =238
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483609
            Name ="FormHeader"
            AutoHeight =1
        End
        Begin Section
            Height =3810
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =360
                    Width =6120
                    Height =330
                    Name ="contactID"
                    ControlSource ="contactID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =360
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =2220
                            Height =330
                            Name ="contactID_Label"
                            Caption ="contactID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =690
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =768
                    Width =6120
                    Height =600
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="contactAddress2"
                    ControlSource ="contactAddress2"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =768
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =1368
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
                            Left =360
                            Top =768
                            Width =2220
                            Height =600
                            Name ="contactAddress2_Label"
                            Caption ="Address2"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =768
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1368
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =1440
                    Width =6120
                    Height =600
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="contactAddress1"
                    ControlSource ="contactAddress1"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =1440
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =2040
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
                            Left =360
                            Top =1440
                            Width =2220
                            Height =600
                            Name ="contactAddress1_Label"
                            Caption ="Address1"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2040
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =2112
                    Width =6120
                    Height =336
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="contactCity"
                    ControlSource ="contactCity"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =2112
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =2448
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
                            Left =360
                            Top =2112
                            Width =2220
                            Height =336
                            Name ="contactCity_Label"
                            Caption ="City"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =2112
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2448
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =2520
                    Width =6120
                    Height =330
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="contactcounty"
                    ControlSource ="contactcounty"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =2850
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =2520
                            Width =2220
                            Height =330
                            Name ="contactcounty_Label"
                            Caption ="contactcounty"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =2520
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2850
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =2652
                    Top =2928
                    Width =6120
                    Height =336
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY [StateAbbrev]; "
                    ColumnWidths ="1440"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =2652
                    LayoutCachedTop =2928
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =3264
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =2928
                            Width =2220
                            Height =336
                            Name ="stateRef_Label"
                            Caption ="State"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =2928
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3264
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2652
                    Top =3336
                    Width =6120
                    Height =336
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="contactZip"
                    ControlSource ="contactZip"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2652
                    LayoutCachedTop =3336
                    LayoutCachedWidth =8772
                    LayoutCachedHeight =3672
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =3336
                            Width =2220
                            Height =336
                            Name ="contactZip_Label"
                            Caption ="Zip Code"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =3336
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3672
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
