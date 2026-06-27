Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
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
    Width =15120
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =3765
    Top =9765
    Right =17895
    Bottom =12285
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6dad0834ed79e440
    End
    RecordSource ="SELECT ATFRevT.* FROM ATFRevT WHERE (((ATFRevT.schref)=[forms]![Schedule_Form]!["
        "schid])) ORDER BY ATFRevT.RevNo DESC;"
    DatasheetFontName ="Calibri"
    Moveable =0
    FilterOnLoad =255
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
            Height =1620
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =8550
                    Top =240
                    Width =255
                    Height =15
                    FontSize =9
                    ForeColor =5026082
                    Name ="Revid"
                    ControlSource ="Revid"

                    LayoutCachedLeft =8550
                    LayoutCachedTop =240
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =4305
                    Top =540
                    Width =255
                    Height =15
                    FontSize =9
                    TabIndex =1
                    ForeColor =5026082
                    Name ="schref"
                    ControlSource ="schref"

                    LayoutCachedLeft =4305
                    LayoutCachedTop =540
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1350
                    Width =405
                    Height =315
                    TabIndex =2
                    Name ="RevNo"
                    ControlSource ="RevNo"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision No required"
                    DefaultValue ="0"

                    LayoutCachedLeft =1350
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1260
                            Height =315
                            Name ="Label2"
                            Caption ="ATF Rev. No:"
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2730
                    Width =1080
                    Height =315
                    TabIndex =3
                    Name ="RevDt"
                    ControlSource ="RevDt"
                    Format ="mm/dd/yy"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"

                    LayoutCachedLeft =2730
                    LayoutCachedWidth =3810
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1860
                            Width =735
                            Height =315
                            Name ="Label3"
                            Caption ="Rev Dt:"
                            LayoutCachedLeft =1860
                            LayoutCachedWidth =2595
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5745
                    Top =15
                    Width =7935
                    Height =1503
                    TabIndex =4
                    Name ="Comment"
                    ControlSource ="Comment"

                    LayoutCachedLeft =5745
                    LayoutCachedTop =15
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =1518
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5295
                            Top =15
                            Width =405
                            Height =315
                            Name ="Label4"
                            Caption ="Nt:"
                            LayoutCachedLeft =5295
                            LayoutCachedTop =15
                            LayoutCachedWidth =5700
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2430
                    Top =510
                    Width =1080
                    Height =315
                    TabIndex =5
                    Name ="ReviewDt"
                    ControlSource ="ReviewDt"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =2430
                    LayoutCachedTop =510
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =825
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1230
                            Top =510
                            Width =1080
                            Height =315
                            Name ="Label5"
                            Caption ="Review Dt:"
                            LayoutCachedLeft =1230
                            LayoutCachedTop =510
                            LayoutCachedWidth =2310
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4680
                    Top =525
                    Width =705
                    Height =300
                    TabIndex =6
                    Name ="reviewer"
                    ControlSource ="reviewer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.OBSOLETE)=No) AND ((Emps."
                        "[LO Exec])=No)) ORDER BY Emps.EmpInt;"
                    Format =">"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =525
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =825
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3600
                            Top =525
                            Width =1020
                            Height =300
                            Name ="Label6"
                            Caption ="Reviewer:"
                            LayoutCachedLeft =3600
                            LayoutCachedTop =525
                            LayoutCachedWidth =4620
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4455
                    Width =780
                    Height =300
                    TabIndex =7
                    Name ="Combo9"
                    ControlSource ="NtNaker"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.OBSOLETE)=No) AND ((Emps."
                        "[LO Exec])=No)) ORDER BY Emps.EmpInt;"
                    ValidationRule ="Is Not Null"
                    ValidationText ="NT By is required (even for 1st ATF)"
                    Format =">"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4455
                    LayoutCachedWidth =5235
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3795
                            Width =615
                            Height =300
                            Name ="Label10"
                            Caption ="Nt By:"
                            LayoutCachedLeft =3795
                            LayoutCachedWidth =4410
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Top =360
                    Width =570
                    Height =315
                    FontSize =9
                    TabIndex =8
                    ForeColor =5026082
                    Name ="Text11"
                    ControlSource ="Revid"

                    LayoutCachedTop =360
                    LayoutCachedWidth =570
                    LayoutCachedHeight =675
                End
            End
        End
    End
End
