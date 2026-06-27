Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    RecordLocks =2
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10380
    DatasheetFontHeight =10
    ItemSuffix =127
    Left =4260
    Top =240
    Right =14310
    Bottom =7425
    RecSrcDt = Begin
        0xa0c37ce9c122e340
    End
    Caption ="Search for Text ..."
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
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
            ForeColor =-2147483630
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
            BorderLineStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =180
                    Top =5100
                    Width =2280
                    Height =300
                    TabIndex =7
                    Name ="btnStartSearch"
                    Caption ="&Start Search"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =5100
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =5400
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =85
                    Left =180
                    Top =5940
                    Width =2280
                    Height =300
                    TabIndex =9
                    Name ="btnClearInputs"
                    Caption ="Clear Inputs/O&utputs"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =5940
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =6240
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    Left =180
                    Top =5460
                    Width =2280
                    Height =300
                    TabIndex =8
                    Name ="btnStopSearch"
                    Caption ="Sto&p Search"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Stops search at end of current search item (ie. at the end of the current table,"
                        " form, etc ...)"
                    UnicodeAccessKey =112

                    LayoutCachedLeft =180
                    LayoutCachedTop =5460
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =5760
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =6660
                    Width =2280
                    Height =300
                    TabIndex =11
                    Name ="btnAbout"
                    Caption ="About"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =6660
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =6960
                End
                Begin Tab
                    OverlapFlags =85
                    BackStyle =0
                    Left =2580
                    Top =600
                    Width =7770
                    Height =6600
                    TabIndex =12
                    Name ="tabResultsPages"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =600
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =7200
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgQueryResults"
                            Caption ="Queries"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =215
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtQueryResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgFormResults"
                            Caption ="Forms"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtFormResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgReportResults"
                            Caption ="Reports"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtReportResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgTableResults"
                            Caption ="Tables"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtTableResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgModuleResults"
                            Caption ="Modules"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtModuleResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgMacroResults"
                            Caption ="Macros"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtMacroResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgC_Bars"
                            Caption ="Command Bars"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    AccessKey =84
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    Name ="txtC_BarResults"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgReplace"
                            Caption ="Edit && Replace"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    ScrollBars =2
                                    OverlapFlags =247
                                    Left =2715
                                    Top =2265
                                    Width =7440
                                    Height =4140
                                    FontSize =10
                                    Name ="txtValue"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =2265
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =6405
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2715
                                            Top =2025
                                            Width =1140
                                            Height =240
                                            Name ="Label125"
                                            Caption ="Value: "
                                            LayoutCachedLeft =2715
                                            LayoutCachedTop =2025
                                            LayoutCachedWidth =3855
                                            LayoutCachedHeight =2265
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =9555
                                    Top =6645
                                    Width =240
                                    TabIndex =1
                                    Name ="chkContinue"
                                    DefaultValue ="False"

                                    LayoutCachedLeft =9555
                                    LayoutCachedTop =6645
                                    LayoutCachedWidth =9795
                                    LayoutCachedHeight =6885
                                    Begin
                                        Begin Label
                                            SpecialEffect =1
                                            BackStyle =1
                                            OverlapFlags =247
                                            Left =8475
                                            Top =6585
                                            Width =1320
                                            Height =300
                                            FontSize =10
                                            BackColor =-2147483633
                                            Name ="Label107"
                                            Caption ="Continue >>"
                                            LayoutCachedLeft =8475
                                            LayoutCachedTop =6585
                                            LayoutCachedWidth =9795
                                            LayoutCachedHeight =6885
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =6195
                                    Top =6675
                                    Width =300
                                    Height =180
                                    TabIndex =2
                                    Name ="chkReplace"
                                    DefaultValue ="False"

                                    LayoutCachedLeft =6195
                                    LayoutCachedTop =6675
                                    LayoutCachedWidth =6495
                                    LayoutCachedHeight =6855
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6420
                                            Top =6645
                                            Width =1800
                                            Height =240
                                            Name ="Label111"
                                            Caption ="Replace with text above"
                                            LayoutCachedLeft =6420
                                            LayoutCachedTop =6645
                                            LayoutCachedWidth =8220
                                            LayoutCachedHeight =6885
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2715
                                    Top =1245
                                    Width =7440
                                    Height =660
                                    TabIndex =3
                                    Name ="txtWhereFound"

                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1245
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =1905
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2715
                                            Top =1005
                                            Width =1140
                                            Height =240
                                            Name ="Label117"
                                            Caption ="Found Where: "
                                            LayoutCachedLeft =2715
                                            LayoutCachedTop =1005
                                            LayoutCachedWidth =3855
                                            LayoutCachedHeight =1245
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2715
                            Top =1005
                            Width =7500
                            Height =6060
                            Name ="pgAbout"
                            Caption ="About"
                            LayoutCachedLeft =2715
                            LayoutCachedTop =1005
                            LayoutCachedWidth =10215
                            LayoutCachedHeight =7065
                            Begin
                                Begin Label
                                    SpecialEffect =2
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =2715
                                    Top =1005
                                    Width =7440
                                    Height =6060
                                    FontSize =10
                                    Name ="lblAbout"
                                    Caption ="Version: 4.2 - For Access 2000 - 2010\015\012Date: 6/2012\015\012\015\012Author:"
                                        " Brent Spaulding\015\012Acknowledgements: Command bar searching was done by some"
                                        "one other than me, the name escapes me, but my gratitude is extended.\015\012\015"
                                        "\012\015\012Note:  The code in this form uses a DAO reference.  The forms Open e"
                                        "vent\015\012           attempts to find this reference.  If it can't be found a "
                                        "message is displayed\015\012           to the user indicating this requirement.\015"
                                        "\012           \015\012           In addition, once this form is imported into y"
                                        "our Access application. You\015\012           should compile your VBA project be"
                                        "cause the code behind this form\015\012           utilizes conditional compilati"
                                        "on to establish connections to a Windows\015\012           API library.\015\012\015"
                                        "\012           If you have problems, contact me (see code for how to)."
                                    LayoutCachedLeft =2715
                                    LayoutCachedTop =1005
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7065
                                End
                            End
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =360
                    Top =1350
                    TabIndex =1
                    Name ="chkQueries"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =590
                            Top =1320
                            Width =795
                            Height =240
                            Name ="Label79"
                            Caption ="Queries"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =1080
                    Width =2280
                    Height =2400
                    Name ="Box80"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =3480
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =300
                    Top =960
                    Width =1110
                    Height =240
                    BackColor =-2147483633
                    Name ="Label81"
                    Caption ="Search What"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =1650
                    TabIndex =2
                    Name ="chkForms"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =590
                            Top =1620
                            Width =795
                            Height =240
                            Name ="Label84"
                            Caption ="Forms"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =1950
                    TabIndex =3
                    Name ="chkReports"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =590
                            Top =1920
                            Width =795
                            Height =240
                            Name ="Label86"
                            Caption ="Reports"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =2250
                    TabIndex =4
                    Name ="chkTables"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =590
                            Top =2220
                            Width =795
                            Height =240
                            Name ="Label88"
                            Caption ="Tables"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =2640
                    Top =240
                    Width =7620
                    Name ="drpSearchText"
                    RowSourceType ="Value List"

                    LayoutCachedLeft =2640
                    LayoutCachedTop =240
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1200
                            Top =240
                            Width =1260
                            Height =240
                            Name ="Label1"
                            Caption ="Search for text:"
                            LayoutCachedLeft =1200
                            LayoutCachedTop =240
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =72
                    Left =180
                    Top =6300
                    Width =2280
                    Height =300
                    TabIndex =10
                    Name ="btnClearHistory"
                    Caption ="Clear Search &History"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =6300
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =6600
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =2550
                    TabIndex =5
                    Name ="chkModules"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =585
                            Top =2520
                            Width =795
                            Height =240
                            Name ="Label91"
                            Caption ="Modules"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1680
                    Top =600
                    Width =780
                    Height =240
                    Name ="Label3"
                    Caption ="Results:"
                    LayoutCachedLeft =1680
                    LayoutCachedTop =600
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =840
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =2850
                    Width =180
                    TabIndex =6
                    Name ="chkMacros"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =590
                            Top =2820
                            Width =810
                            Height =240
                            Name ="Label120"
                            Caption ="Macros"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =360
                    Top =3150
                    Width =180
                    TabIndex =13
                    Name ="chkCBars"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =585
                            Top =3120
                            Width =1140
                            Height =240
                            Name ="Label122"
                            Caption ="Command Bars"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =360
                    Top =3870
                    Width =420
                    Height =210
                    TabIndex =14
                    Name ="chkSkipReplacement"
                    DefaultValue ="True"
                    ControlTipText ="If disabled, you are given the opportuntity to modify property values when then "
                        "contain the pattern being searched.  This is implemented for Queries, Forms, and"
                        " Reports and the object contained with in them."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =600
                            Top =3840
                            Width =1800
                            Height =240
                            Name ="Label113"
                            Caption ="Do Not Offer to Replace"
                            LayoutCachedLeft =600
                            LayoutCachedTop =3840
                            LayoutCachedWidth =2400
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =3660
                    Width =2280
                    Height =1320
                    Name ="Box123"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3660
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =4980
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =300
                    Top =3540
                    Width =1305
                    Height =240
                    BackColor =-2147483633
                    Name ="Label124"
                    Caption ="Edit and Replace"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =300
                    Top =4140
                    Width =2100
                    Height =795
                    Name ="Label126"
                    Caption ="(Replace is for Property values of Queries, Forms, Reports, and the objects cont"
                        "ained with in them)"
                    LayoutCachedLeft =300
                    LayoutCachedTop =4140
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =4935
                End
            End
        End
    End
End
CodeBehindForm
' See "xdlgSearchForText_v42.cls"
