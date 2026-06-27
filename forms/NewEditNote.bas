Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14565
    DatasheetFontHeight =11
    Left =4455
    Top =4155
    Right =20910
    Bottom =12870
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x9089e3caf156e440
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
            Height =255
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483633
                    BorderColor =16777215
                    Name ="noteDateTime_Label"
                    Caption ="Date/Time"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1980
                    Width =690
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =-2147483633
                    BorderColor =16777215
                    Name ="noteMemo_Label"
                    Caption ="Details"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1980
                    LayoutCachedWidth =2670
                    LayoutCachedHeight =255
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =2835
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =105
                    Top =2040
                    Width =540
                    FontSize =8
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =1643706
                    Name ="noteID"
                    ControlSource ="noteID"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =105
                    LayoutCachedTop =2040
                    LayoutCachedWidth =645
                    LayoutCachedHeight =2280
                End
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =15
                    Top =1425
                    Width =14415
                    Height =1260
                    FontSize =8
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =11830108
                    ForeColor =-2147483640
                    Name ="noteMemo"
                    ControlSource ="noteMemo"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =15
                    LayoutCachedTop =1425
                    LayoutCachedWidth =14430
                    LayoutCachedHeight =2685
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =105
                    Top =1755
                    Width =450
                    Height =255
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =1643706
                    Name ="ProposalRef"
                    ControlSource ="groupRef"
                    StatusBarText ="link to Proposal table"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =105
                    LayoutCachedTop =1755
                    LayoutCachedWidth =555
                    LayoutCachedHeight =2010
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =105
                    Top =2325
                    Width =465
                    Height =255
                    FontSize =8
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =1643706
                    Name ="brokerRef"
                    ControlSource ="dealRef"
                    StatusBarText ="link to broker table"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =105
                    LayoutCachedTop =2325
                    LayoutCachedWidth =570
                    LayoutCachedHeight =2580
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =105
                    Top =1515
                    Width =465
                    Height =255
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =1643706
                    Name ="clientRef"
                    ControlSource ="clientRef"
                    StatusBarText ="link to client table"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =105
                    LayoutCachedTop =1515
                    LayoutCachedWidth =570
                    LayoutCachedHeight =1770
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1515
                    Top =735
                    Width =1350
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label25"
                    Caption ="FA Attendees:"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =1515
                    LayoutCachedTop =735
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3815
                    Top =1125
                    Width =420
                    Height =255
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="NoTargetAttendees"
                    ControlSource ="NoTargetAttendees"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3815
                    LayoutCachedTop =1125
                    LayoutCachedWidth =4235
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1470
                            Top =1125
                            Width =2220
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label26"
                            Caption ="Total Client/Target Attendees"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1470
                            LayoutCachedTop =1125
                            LayoutCachedWidth =3690
                            LayoutCachedHeight =1365
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11150
                    Top =660
                    Width =480
                    Height =255
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="NoOEM"
                    ControlSource ="NoOEM"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Number of Attendees"

                    LayoutCachedLeft =11150
                    LayoutCachedTop =660
                    LayoutCachedWidth =11630
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =10545
                            Top =660
                            Width =540
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label28"
                            Caption ="OEM:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =10545
                            LayoutCachedTop =660
                            LayoutCachedWidth =11085
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =12645
                    Top =660
                    Width =540
                    Height =255
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="NoPeopleNet"
                    ControlSource ="NoPeopleNet"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Number of Attendees"

                    LayoutCachedLeft =12645
                    LayoutCachedTop =660
                    LayoutCachedWidth =13185
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =11680
                            Top =660
                            Width =900
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label29"
                            Caption ="PeopleNet:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =11680
                            LayoutCachedTop =660
                            LayoutCachedWidth =12580
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13855
                    Top =660
                    Width =540
                    Height =255
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="NoOther"
                    ControlSource ="NoOther"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Number of Attendees"

                    LayoutCachedLeft =13855
                    LayoutCachedTop =660
                    LayoutCachedWidth =14395
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =13250
                            Top =660
                            Width =540
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label30"
                            Caption ="Other:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =13250
                            LayoutCachedTop =660
                            LayoutCachedWidth =13790
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =12
                    ListWidth =2880
                    Left =5700
                    Width =900
                    Height =255
                    FontSize =8
                    TabIndex =9
                    BackColor =11271167
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="noteUser"
                    ControlSource ="noteUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm, Emps.OBSOLETE FROM Emps WHERE (((Emps.OBSOLETE)="
                        "No)) ORDER BY Emps.EmpInt, Emps.[LO Exec];"
                    ColumnWidths ="720;2160"
                    StatusBarText ="entered by salesman"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Entered by is required"
                    FontName ="MS Sans Serif"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4800
                            Width =885
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label31"
                            Caption ="Entered by:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4800
                            LayoutCachedWidth =5685
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13155
                    Top =15
                    Width =1140
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="noteFU"
                    ControlSource ="noteFU"
                    Format ="mm/dd/yy"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =13155
                    LayoutCachedTop =15
                    LayoutCachedWidth =14295
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12090
                            Top =15
                            Width =1035
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label33"
                            Caption ="Follow Up Dt:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =12090
                            LayoutCachedTop =15
                            LayoutCachedWidth =13125
                            LayoutCachedHeight =255
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Width =1770
                    Height =255
                    FontSize =8
                    TabIndex =11
                    BackColor =11271167
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="CallType"
                    ControlSource ="CallType"
                    RowSourceType ="Table/Query"
                    RowSource ="NoteType"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Call type is required"
                    FontName ="MS Sans Serif"
                    Format =">"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =840
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =735
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label36"
                            Caption ="CallType:"
                            FontName ="MS Sans Serif"
                            LayoutCachedWidth =735
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3405
                    Top =720
                    TabIndex =12
                    Name ="JF"
                    ControlSource ="JF"
                    StatusBarText ="attended"

                    LayoutCachedLeft =3405
                    LayoutCachedTop =720
                    LayoutCachedWidth =3665
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3730
                            Top =720
                            Width =240
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label42"
                            Caption ="JF"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3730
                            LayoutCachedTop =720
                            LayoutCachedWidth =3970
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =4560
                    Top =720
                    Width =395
                    TabIndex =13
                    Name ="PF"
                    ControlSource ="PF"
                    StatusBarText ="attended"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =720
                    LayoutCachedWidth =4955
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =4885
                            Top =720
                            Width =330
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label44"
                            Caption ="PF"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4885
                            LayoutCachedTop =720
                            LayoutCachedWidth =5215
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3675
                    Top =15
                    Width =1020
                    FontSize =8
                    TabIndex =14
                    BackColor =11271167
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="NoteEctDt"
                    ControlSource ="NoteEctDt"
                    Format ="mm/dd/yy"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Actual date is required"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =15
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2715
                            Top =15
                            Width =930
                            Height =240
                            FontSize =7
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =255
                            Name ="Label50"
                            Caption ="Actual Dt:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =15
                            LayoutCachedWidth =3645
                            LayoutCachedHeight =255
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4425
                    Top =345
                    TabIndex =15
                    Name ="OnSiteResch"
                    ControlSource ="OnSiteResch"
                    StatusBarText ="on-site meeting rescheduled?"

                    LayoutCachedLeft =4425
                    LayoutCachedTop =345
                    LayoutCachedWidth =4685
                    LayoutCachedHeight =585
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4665
                            Top =345
                            Width =1365
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label53"
                            Caption ="ReScheduled?"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4665
                            LayoutCachedTop =345
                            LayoutCachedWidth =6030
                            LayoutCachedHeight =585
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1545
                    Top =345
                    Width =900
                    FontSize =8
                    TabIndex =16
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="OnSiteSched"
                    ControlSource ="OnSiteSched"
                    Format ="mm/dd/yy"
                    StatusBarText ="date on onsite meeting is scheduled"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =1545
                    LayoutCachedTop =345
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =585
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =210
                            Top =345
                            Width =1170
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label54"
                            Caption ="On-Site Sch:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =210
                            LayoutCachedTop =345
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =585
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2865
                    Top =345
                    TabIndex =17
                    Name ="OnSiteXCL"
                    ControlSource ="OnSiteXCL"
                    StatusBarText ="scheduled onsite meeting cancelled?"

                    LayoutCachedLeft =2865
                    LayoutCachedTop =345
                    LayoutCachedWidth =3125
                    LayoutCachedHeight =585
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3090
                            Top =345
                            Width =1170
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label57"
                            Caption ="Cancelled?"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3090
                            LayoutCachedTop =345
                            LayoutCachedWidth =4260
                            LayoutCachedHeight =585
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    ScrollBarAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =7500
                    Width =4320
                    Height =255
                    FontSize =8
                    TabIndex =18
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="ContactNM"
                    ControlSource ="ContactNM"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.contactFullName, Contacts.contactTitle FROM Contacts WHERE (((Co"
                        "ntacts.clientRef)=[forms]![ClientsView]![clientid]) AND ((Contacts.ContactObsole"
                        "te)=No)) ORDER BY Contacts.clientRef, Contacts.contactLast, Contacts.contactFirs"
                        "t;"
                    ColumnWidths ="2160;1440"
                    StatusBarText ="link to contacts table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Contact is Required. Use Not Applicable for Note Only\""
                    DefaultValue ="\"Contact required if not Note\""
                    FontName ="MS Sans Serif"
                    CanGrow =255
                    ListItemsEditForm ="Contacts_viewadd"
                    InheritValueList =0

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6705
                            Width =735
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label59"
                            Caption ="Contact:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6705
                            LayoutCachedWidth =7440
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =1140
                    Width =1065
                    FontSize =8
                    TabIndex =19
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="noteDateTime"
                    ControlSource ="noteDateTime"
                    Format ="mm/dd/yy"
                    FontName ="MS Sans Serif"

                    LayoutCachedTop =1140
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =1380
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =8010
                    Top =300
                    Width =2565
                    Height =210
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label63"
                    Caption ="Double click to add new contact"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =8010
                    LayoutCachedTop =300
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =510
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7740
                    Top =1125
                    Width =3165
                    Height =255
                    FontSize =8
                    TabIndex =20
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="Other Location"
                    ControlSource ="Other Location"
                    StatusBarText ="25"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Other_Location"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =1125
                    LayoutCachedWidth =10905
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4530
                            Top =1125
                            Width =3030
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label64"
                            Caption ="Meeting Location (if not contact office):"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4530
                            LayoutCachedTop =1125
                            LayoutCachedWidth =7560
                            LayoutCachedHeight =1365
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =3210
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box65"
                    LayoutCachedLeft =3210
                    LayoutCachedTop =675
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =1050
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =5655
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box67"
                    LayoutCachedLeft =5655
                    LayoutCachedTop =675
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1050
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =10425
                    Top =630
                    Width =4140
                    Height =375
                    Name ="Box70"
                    LayoutCachedLeft =10425
                    LayoutCachedTop =630
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =1005
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =7080
                    Top =690
                    TabIndex =21
                    Name ="TC"
                    ControlSource ="TC"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =690
                    LayoutCachedWidth =7340
                    LayoutCachedHeight =930
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =7310
                            Top =690
                            Width =330
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label74"
                            Caption ="TC"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =7310
                            LayoutCachedTop =690
                            LayoutCachedWidth =7640
                            LayoutCachedHeight =930
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =6885
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box75"
                    LayoutCachedLeft =6885
                    LayoutCachedTop =675
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =1050
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5805
                    Top =720
                    TabIndex =22
                    Name ="MM"
                    ControlSource ="MM"
                    StatusBarText ="Mike Meehan added 1-11-13"

                    LayoutCachedLeft =5805
                    LayoutCachedTop =720
                    LayoutCachedWidth =6065
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6035
                            Top =720
                            Width =345
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label76"
                            Caption ="MM"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6035
                            LayoutCachedTop =720
                            LayoutCachedWidth =6380
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =4440
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box77"
                    LayoutCachedLeft =4440
                    LayoutCachedTop =675
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =1050
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9180
                    Top =720
                    TabIndex =23
                    Name ="Other"
                    ControlSource ="Other"
                    StatusBarText ="other FA person at meeting"

                    LayoutCachedLeft =9180
                    LayoutCachedTop =720
                    LayoutCachedWidth =9440
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9410
                            Top =720
                            Width =555
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label79"
                            Caption ="Other"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =9410
                            LayoutCachedTop =720
                            LayoutCachedWidth =9965
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =9090
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box80"
                    LayoutCachedLeft =9090
                    LayoutCachedTop =675
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =1050
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8205
                    Top =675
                    TabIndex =24
                    Name ="MS"
                    ControlSource ="MS"

                    LayoutCachedLeft =8205
                    LayoutCachedTop =675
                    LayoutCachedWidth =8465
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =8430
                            Top =675
                            Width =360
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label82"
                            Caption ="MS"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =8430
                            LayoutCachedTop =675
                            LayoutCachedWidth =8790
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =8010
                    Top =675
                    Width =945
                    Height =375
                    Name ="Box83"
                    LayoutCachedLeft =8010
                    LayoutCachedTop =675
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =1050
                End
            End
        End
        Begin FormFooter
            Height =45
            Name ="FormFooter"
        End
    End
End
