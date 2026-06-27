Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10845
    DatasheetFontHeight =11
    ItemSuffix =39
    Left =3885
    Top =2775
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x15ae20fb0934e440
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin PageHeader
            Height =2610
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =105
                    Top =135
                    Width =10485
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="BP Products North America Inc./Fleet Advantage, LLC"
                    LayoutCachedLeft =105
                    LayoutCachedTop =135
                    LayoutCachedWidth =10590
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =555
                    Width =10455
                    Height =405
                    FontSize =14
                    FontWeight =700
                    Name ="Text2"
                    ControlSource ="=\"Lease Schedule Number \" & [Forms]![Schedule_Form]![schno]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =555
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =960
                End
                Begin Label
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =120
                    Top =1185
                    Width =10455
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Attachment 2"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1185
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =1545
                End
                Begin Label
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =120
                    Top =1440
                    Width =10455
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Equipment Acceptance Notification Certificate"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =1800
                End
                Begin Label
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =120
                    Top =2220
                    Width =10455
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label8"
                    Caption ="(To be signed and returned to Lessor after Lessee's inspection of the delivered "
                        "items(s) of Equipment)"
                    LayoutCachedLeft =120
                    LayoutCachedTop =2220
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =2580
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =1755
                    Width =10380
                    Height =390
                    FontSize =18
                    TabIndex =1
                    ForeColor =255
                    Name ="Text33"
                    ControlSource ="=IIf(IsNull([Forms]![Schedule_Form]![SchedDt]),\"ERROR - NO Schedule Date\",Null"
                        ")"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1755
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =2145
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =210
                    Top =780
                    Width =10380
                    Height =390
                    FontSize =18
                    TabIndex =2
                    ForeColor =255
                    Name ="Text35"
                    ControlSource ="=IIf(IsNull([Forms]![Schedule_Form]![MLdt]),\"ERROR - NO Master Lease Date\",Nul"
                        "l)"

                    LayoutCachedLeft =210
                    LayoutCachedTop =780
                    LayoutCachedWidth =10590
                    LayoutCachedHeight =1170
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =9480
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =195
                    Width =10410
                    Height =1905
                    Name ="Text9"
                    ControlSource ="=\"Reference is made to that certain Master Lease Agreement as amended, restated"
                        ", supplemented or otherwise  modified from time to time, the \"\"Master Lease\"\""
                        ") dated \" & MonthName(DatePart(\"m\",[Forms]![Schedule_Form]![MLDT])) & \" \" &"
                        " DatePart(\"d\",[Forms]![Schedule_Form]![MLDT]) & \", \" & DatePart(\"yyyy\",[Fo"
                        "rms]![Schedule_Form]![MLDT]) & \", between Fleet Advantage, LLC  (together with "
                        "its permitted assigns, \"\"Lessor \"\")  and  BP Products North America Inc. (\""
                        "\"Lessee \"\") and that certain Lease Schedule Number \" & [Forms]![Schedule_For"
                        "m]![schno] & \", dated \" & MonthName(DatePart(\"m\",[Forms]![Schedule_Form]![sc"
                        "heddt])) & \" \" & DatePart(\"d\",[Forms]![Schedule_Form]![scheddt]) & \", \" & "
                        "DatePart(\"yyyy\",[Forms]![Schedule_Form]![scheddt]) & \", \" & \"relating to th"
                        "e Master Lease (the \"\"Lease Schedule\"\") \" & \"This Attachment 2 – Equipment"
                        " Acceptance Notification Certificate is attached thereto and made a part thereof"
                        ".  Capitalized terms used but not defined herein shall have the meanings given s"
                        "uch terms in the Master Lease or Lease Schedule, as applicable.\""

                    LayoutCachedLeft =195
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =1905
                End
                Begin Label
                    Left =180
                    Top =1950
                    Width =10440
                    Height =2235
                    Name ="Label11"
                    Caption ="Lessee certifies that the Equipment , described on Exhibit A attached hereto,  l"
                        "eased under the Lease Agreement and Lease Schedule to which this Attachment 2 is"
                        " attached, was delivered in good order and condition, installed, tested and is a"
                        "cceptable to us, and is in use as of the date hereof.  The Equipment is suitable"
                        " for purposes of accepting it under the Master Lease and Lease Schedule and maki"
                        "ng it subject to the Master Lease and Lease Schedule.  The decals, labels, etc.,"
                        " if required and supplied, have been affixed to the Equipment.  Lessee approves "
                        "payment by Lessor to the manufacturer for the Equipment as well as all services "
                        "described in the Master Lease and Lease Schedule.  Lessee agrees that its paymen"
                        "t obligation to Lessor under the Master Lease and related Lease Schedule has com"
                        "menced. "
                    LayoutCachedLeft =180
                    LayoutCachedTop =1950
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =4185
                End
                Begin Label
                    Left =195
                    Top =4245
                    Width =10440
                    Height =1680
                    Name ="Label12"
                    Caption ="Lessee is aware of, and accepts, its obligations under the Master Lease and rela"
                        "ted Lease Schedule and agrees to enforce in Lessee’s own name all warranties, ag"
                        "reements, or representations, if any, which may be made by the Supplier to Lesse"
                        "e (or assigned to Lessee pursuant to the Master Lease and Lease Schedule).  Less"
                        "ee certifies that the Equipment is available for service as of the date below an"
                        "d agrees that said date shall be deemed to be the Acceptance Date under the Mast"
                        "er Lease and Lease Schedule. Lessee also certifies that the Equipment will be us"
                        "ed for lawful business purposes only."
                    LayoutCachedLeft =195
                    LayoutCachedTop =4245
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =5925
                End
                Begin Label
                    Left =195
                    Top =5985
                    Width =10440
                    Height =1365
                    Name ="Label13"
                    Caption ="Lessee understands and agrees that Per Diem Rent shall not begin to accrue until"
                        " the In-Service Acceptance Date which shall be the 60th day following after the "
                        "Acceptance Date as described under Vehicle In-Service and Set-Up Period set fort"
                        "h on Attachment 1 – Additional Terms to the Lease Schedule, and the Rental Payme"
                        "nts shall begin to accrue on the Base Lease Commencement Date in accordance with"
                        " the Master Lease and related Lease Schedule."
                    LayoutCachedLeft =195
                    LayoutCachedTop =5985
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =7350
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =195
                    Top =7410
                    Width =10245
                    Height =555
                    TabIndex =1
                    Name ="Text14"
                    ControlSource ="=\"This Equipment Acceptance Notification Certificate is hereby executed and agr"
                        "eed as of \" & \"_________________\" & \" (the \"\"Acceptance Date\"\") by Lesse"
                        "e.\""

                    LayoutCachedLeft =195
                    LayoutCachedTop =7410
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =7965
                End
                Begin Label
                    Left =270
                    Top =8250
                    Width =795
                    Height =330
                    Name ="Label16"
                    Caption ="Lessee:"
                    LayoutCachedLeft =270
                    LayoutCachedTop =8250
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =8580
                End
                Begin Label
                    Left =270
                    Top =8700
                    Width =795
                    Height =330
                    Name ="Label17"
                    Caption ="Name:"
                    LayoutCachedLeft =270
                    LayoutCachedTop =8700
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =9030
                End
                Begin Label
                    Left =270
                    Top =9150
                    Width =795
                    Height =330
                    Name ="Label18"
                    Caption ="Title:"
                    LayoutCachedLeft =270
                    LayoutCachedTop =9150
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =9480
                End
                Begin Line
                    Left =1170
                    Top =8505
                    Width =6045
                    Name ="Line19"
                    LayoutCachedLeft =1170
                    LayoutCachedTop =8505
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =8505
                End
                Begin Line
                    Left =1170
                    Top =8925
                    Width =6045
                    Name ="Line20"
                    LayoutCachedLeft =1170
                    LayoutCachedTop =8925
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =8925
                End
                Begin Line
                    Left =1170
                    Top =9330
                    Width =6045
                    Name ="Line21"
                    LayoutCachedLeft =1170
                    LayoutCachedTop =9330
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =9330
                End
            End
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
        End
    End
End
