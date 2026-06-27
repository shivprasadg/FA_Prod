Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13980
    DatasheetFontHeight =10
    ItemSuffix =119
    Left =2970
    Top =4695
    Right =16965
    Bottom =12060
    AfterDelConfirm ="[Event Procedure]"
    Filter ="((SpecReviews.SRSpecDate Is Not Null)) And (SpecReviews.SRSpecDate Is Not Null)"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xe94aa222f383e340
    End
    RecordSource ="SpecReviews"
    Caption ="Spec Review"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnGotFocus ="Refresh"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Spec Reviews"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
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
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =330
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9075
                    Width =420
                    Height =255
                    ForeColor =2366701
                    Name ="groupRef"
                    ControlSource ="groupRef"
                    StatusBarText ="link to groups table"

                    LayoutCachedLeft =9075
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5850
                    Width =2445
                    Height =255
                    ColumnWidth =1815
                    TabIndex =1
                    Name ="SRDateAdded"
                    ControlSource ="SRDateAdded"
                    StatusBarText ="this is the date the record was added to the db"

                    LayoutCachedLeft =5850
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =4830
                            Width =990
                            Height =240
                            ForeColor =16777215
                            Name ="Label95"
                            Caption ="Date Added:"
                            LayoutCachedLeft =4830
                            LayoutCachedWidth =5820
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9615
                    Width =555
                    Height =255
                    TabIndex =2
                    ForeColor =2366701
                    Name ="SRID"
                    ControlSource ="SRID"

                    LayoutCachedLeft =9615
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =255
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =30
                    Width =1650
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    ForeColor =-2147483615
                    Name ="SaveSpecBtn"
                    Caption ="Add Spec Review"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Save Record"

                    LayoutCachedLeft =30
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5265
            Name ="Detail"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =9540
                    Top =2280
                    Width =4440
                    Height =2940
                    BackColor =14545355
                    Name ="Box118"
                    LayoutCachedLeft =9540
                    LayoutCachedTop =2280
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =5220
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Top =360
                    Width =9495
                    Height =4875
                    BackColor =11853709
                    Name ="Box89"
                    LayoutCachedTop =360
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =5235
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =9540
                    Top =360
                    Width =4440
                    Height =1875
                    BackColor =13298606
                    Name ="Box90"
                    LayoutCachedLeft =9540
                    LayoutCachedTop =360
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =2235
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1260
                    Top =750
                    Width =2160
                    BorderColor =13017476
                    Name ="SRWB"
                    ControlSource ="SRWB"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =750
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =990
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =60
                            Top =750
                            Width =1095
                            Height =240
                            Name ="Label3"
                            Caption ="WB:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =750
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1260
                    Top =1050
                    Width =2160
                    TabIndex =2
                    BorderColor =13017476
                    Name ="SREngine"
                    ControlSource ="SREngine"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1050
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1290
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =60
                            Top =1050
                            Width =1095
                            Height =240
                            Name ="Label4"
                            Caption ="Engine OEM:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1050
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =1290
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =1950
                    Width =2160
                    TabIndex =6
                    BorderColor =13017476
                    Name ="SRTransmission"
                    ControlSource ="SRTransmission"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =1950
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =1950
                            Width =1095
                            Height =240
                            Name ="Label5"
                            Caption ="Transmission:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =1950
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =2190
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =2250
                    Width =2160
                    TabIndex =8
                    BorderColor =13017476
                    Name ="SRFrontAxle"
                    ControlSource ="SRFrontAxle"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =2250
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =2250
                            Width =1095
                            Height =240
                            Name ="Label6"
                            Caption ="Front Axle:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =2250
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =2490
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =2550
                    Width =2160
                    TabIndex =10
                    BorderColor =13017476
                    Name ="SRRearAxle"
                    ControlSource ="SRRearAxle"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =2550
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2790
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =2550
                            Width =1095
                            Height =240
                            Name ="Label7"
                            Caption ="Rear Axle:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =2550
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =2790
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =2850
                    Width =2160
                    TabIndex =12
                    BorderColor =13017476
                    Name ="SRRatio"
                    ControlSource ="SRRatio"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =2850
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3090
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =2850
                            Width =1095
                            Height =240
                            Name ="Label8"
                            Caption ="Ratio:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =2850
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =3090
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =3150
                    Width =2160
                    TabIndex =14
                    BorderColor =13017476
                    Name ="SRSuspension"
                    ControlSource ="SRSuspension"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =3150
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3390
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =3150
                            Width =1095
                            Height =240
                            Name ="Label9"
                            Caption ="Suspension:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =3150
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =3390
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =3450
                    Width =2160
                    TabIndex =16
                    BorderColor =13017476
                    Name ="SR5thWheel"
                    ControlSource ="SR5thWheel"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =3450
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3690
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =3450
                            Width =1095
                            Height =240
                            Name ="Label10"
                            Caption ="5th Wheel:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =3450
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =3690
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =3750
                    Width =2160
                    TabIndex =18
                    BorderColor =13017476
                    Name ="SRFrontWheels"
                    ControlSource ="SRFrontWheels"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =3750
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3990
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =3750
                            Width =1095
                            Height =240
                            Name ="Label11"
                            Caption ="Front Wheels:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =3750
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =3990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =4050
                    Width =2160
                    TabIndex =20
                    BorderColor =13017476
                    Name ="SRFrontTires"
                    ControlSource ="SRFrontTires"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =4050
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =4290
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =45
                            Top =4050
                            Width =1095
                            Height =240
                            Name ="Label12"
                            Caption ="Front Tires:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =4050
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =4290
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1230
                    Top =4335
                    Width =2160
                    TabIndex =22
                    BorderColor =13017476
                    Name ="SRRearWheels"
                    ControlSource ="SRRearWheels"

                    LayoutCachedLeft =1230
                    LayoutCachedTop =4335
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =4575
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =30
                            Top =4335
                            Width =1095
                            Height =240
                            Name ="Label13"
                            Caption ="Rear Wheels:"
                            LayoutCachedLeft =30
                            LayoutCachedTop =4335
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =4575
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1230
                    Top =4605
                    Width =2160
                    TabIndex =24
                    BorderColor =13017476
                    Name ="SRRearTires"
                    ControlSource ="SRRearTires"

                    LayoutCachedLeft =1230
                    LayoutCachedTop =4605
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =4845
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =30
                            Top =4605
                            Width =1095
                            Height =240
                            Name ="Label14"
                            Caption ="Rear Tires:"
                            LayoutCachedLeft =30
                            LayoutCachedTop =4605
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =4845
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =1950
                    Width =2160
                    TabIndex =31
                    BorderColor =13017476
                    Name ="SRJake"
                    ControlSource ="SRJake"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1950
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =1950
                            Width =1095
                            Height =240
                            Name ="Label15"
                            Caption ="Jake:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =1950
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =2190
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =2250
                    Width =2160
                    TabIndex =33
                    BorderColor =13017476
                    Name ="SRFuelTanks"
                    ControlSource ="SRFuelTanks"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2250
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =2250
                            Width =1095
                            Height =240
                            Name ="Label16"
                            Caption ="Fuel Tanks:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =2250
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =2490
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =2550
                    Width =2160
                    TabIndex =34
                    BorderColor =13017476
                    Name ="SRWeight"
                    ControlSource ="SRWeight"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2550
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =2790
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =2550
                            Width =1095
                            Height =240
                            Name ="Label17"
                            Caption ="Weight:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =2550
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =2790
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =2850
                    Width =2160
                    TabIndex =35
                    BorderColor =13017476
                    Name ="SRCabColor"
                    ControlSource ="SRCabColor"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2850
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =3090
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =2850
                            Width =1095
                            Height =240
                            Name ="Label18"
                            Caption ="Cab Color:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =2850
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =3090
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3465
                    Top =750
                    Width =1095
                    TabIndex =1
                    BorderColor =13017476
                    Name ="SRWBRebate"
                    ControlSource ="SRWBRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3465
                    LayoutCachedTop =750
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =990
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3465
                    Top =1050
                    Width =1095
                    TabIndex =3
                    BorderColor =13017476
                    Name ="SREngineRebate"
                    ControlSource ="SREngineRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3465
                    LayoutCachedTop =1050
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =1290
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =1950
                    Width =1095
                    TabIndex =7
                    BorderColor =13017476
                    Name ="SRTransmissionRebate"
                    ControlSource ="SRTransmissionRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =1950
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =2190
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =2250
                    Width =1095
                    TabIndex =9
                    BorderColor =13017476
                    Name ="SRFrontAxleRebate"
                    ControlSource ="SRFrontAxleRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =2250
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =2490
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =2550
                    Width =1095
                    TabIndex =11
                    BorderColor =13017476
                    Name ="SRRearAxleRebate"
                    ControlSource ="SRRearAxleRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =2550
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =2790
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =2850
                    Width =1095
                    TabIndex =13
                    BorderColor =13017476
                    Name ="SRRatioRebate"
                    ControlSource ="SRRatioRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =2850
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =3090
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =3150
                    Width =1095
                    TabIndex =15
                    BorderColor =13017476
                    Name ="SRSuspensionRebate"
                    ControlSource ="SRSuspensionRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =3150
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =3390
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =3450
                    Width =1095
                    TabIndex =17
                    BorderColor =13017476
                    Name ="SR5thWheelRebate"
                    ControlSource ="SR5thWheelRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =3450
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =3690
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =3750
                    Width =1095
                    TabIndex =19
                    BorderColor =13017476
                    Name ="SRFrontWheelsRebate"
                    ControlSource ="SRFrontWheelsRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =3750
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =3990
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =4050
                    Width =1095
                    TabIndex =21
                    BorderColor =13017476
                    Name ="SRFrontTiresRebate"
                    ControlSource ="SRFrontTiresRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =4050
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =4290
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =4335
                    Width =1095
                    TabIndex =23
                    BorderColor =13017476
                    Name ="SRRearWheelsRebate"
                    ControlSource ="SRRearWheelsRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =4335
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =4575
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3450
                    Top =4605
                    Width =1095
                    TabIndex =25
                    BorderColor =13017476
                    Name ="SRRearTiresRebate"
                    ControlSource ="SRRearTiresRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3450
                    LayoutCachedTop =4605
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =4845
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =1950
                    Width =1095
                    TabIndex =32
                    BorderColor =13017476
                    Name ="SRJakeRebate"
                    ControlSource ="SRJakeRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =1950
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =2190
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    Left =3480
                    Top =390
                    Width =1050
                    Height =240
                    ForeColor =8421504
                    Name ="Label42"
                    Caption ="Rebate $"
                    LayoutCachedLeft =3480
                    LayoutCachedTop =390
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =630
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    Left =8220
                    Top =390
                    Width =1035
                    Height =240
                    ForeColor =8421504
                    Name ="Label43"
                    Caption ="Rebate $"
                    LayoutCachedLeft =8220
                    LayoutCachedTop =390
                    LayoutCachedWidth =9255
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1260
                    Top =1350
                    Width =960
                    Height =255
                    TabIndex =4
                    BorderColor =13017476
                    Name ="SREngineHP"
                    ControlSource ="SREngineHP"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1350
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1605
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =60
                            Top =1350
                            Width =1095
                            Height =240
                            Name ="Label65"
                            Caption ="Engine HP:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1350
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =1590
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =2280
                    Top =1350
                    Width =270
                    Height =210
                    BorderColor =13017476
                    Name ="Label66"
                    Caption ="HP"
                    LayoutCachedLeft =2280
                    LayoutCachedTop =1350
                    LayoutCachedWidth =2550
                    LayoutCachedHeight =1560
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =6310
                    Top =3150
                    TabIndex =36
                    Name ="SRIntTrim"
                    ControlSource ="SRIntTrim"

                    LayoutCachedLeft =6310
                    LayoutCachedTop =3150
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3390
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =3150
                            Width =1275
                            Height =240
                            Name ="Label68"
                            Caption ="Interior Trim Pkg"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =3150
                            LayoutCachedWidth =6075
                            LayoutCachedHeight =3390
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =4905
                    Width =1095
                    Height =255
                    TabIndex =47
                    BorderColor =13017476
                    Name ="SRMfgRebate"
                    ControlSource ="SRMfgRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =4905
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =5160
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =4905
                            Width =1095
                            Height =240
                            Name ="Label70"
                            Caption ="Mfg Rebate:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =4905
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =5145
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =3435
                    Width =1095
                    Height =255
                    TabIndex =38
                    BorderColor =13017476
                    Name ="SRFilterRebate"
                    ControlSource ="SRFilterRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =3435
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =3690
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =3450
                            Width =1095
                            Height =240
                            Name ="Label71"
                            Caption ="Filter:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =3450
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =3690
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =3735
                    Width =1095
                    Height =255
                    TabIndex =40
                    BorderColor =13017476
                    Name ="SRFanClutchRebate"
                    ControlSource ="SRFanClutchRebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =3735
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =3990
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =3750
                            Width =1095
                            Height =240
                            Name ="Label72"
                            Caption ="Fan Clutch:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =3750
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =3990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =6000
                    Top =4035
                    Width =2160
                    Height =255
                    TabIndex =41
                    BorderColor =13017476
                    Name ="SRMiscRebate1Desc"
                    ControlSource ="SRMiscRebate1Desc"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4035
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =4290
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =4035
                    Width =1095
                    Height =255
                    TabIndex =42
                    BorderColor =13017476
                    Name ="SRMiscRebate1"
                    ControlSource ="SRMiscRebate1"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =4035
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =4290
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =6000
                    Top =4335
                    Width =2160
                    Height =255
                    TabIndex =43
                    BorderColor =13017476
                    Name ="SRMiscRebate2Desc"
                    ControlSource ="SRMiscRebate2Desc"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4335
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =4590
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =4335
                    Width =1095
                    Height =255
                    TabIndex =44
                    BorderColor =13017476
                    Name ="SRMiscRebate2"
                    ControlSource ="SRMiscRebate2"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =4335
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =4590
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =6000
                    Top =4635
                    Width =2160
                    Height =255
                    TabIndex =45
                    BorderColor =13017476
                    Name ="SRMiscRebate3Desc"
                    ControlSource ="SRMiscRebate3Desc"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4635
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =4890
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8340
                    Top =4635
                    Width =1095
                    Height =255
                    TabIndex =46
                    BorderColor =13017476
                    Name ="SRMiscRebate3"
                    ControlSource ="SRMiscRebate3"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =4635
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =4890
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11100
                    Top =750
                    Width =1500
                    Height =255
                    TabIndex =48
                    BorderColor =13017476
                    Name ="TrailerLength"
                    ControlSource ="TrailerLength"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11100
                    LayoutCachedTop =750
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1005
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9780
                            Top =750
                            Width =1260
                            Height =240
                            Name ="Label79"
                            Caption ="Trailer Length:"
                            LayoutCachedLeft =9780
                            LayoutCachedTop =750
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11100
                    Top =1050
                    Width =1500
                    Height =255
                    TabIndex =49
                    BorderColor =13017476
                    Name ="TrailerWidth"
                    ControlSource ="TrailerWidth"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11100
                    LayoutCachedTop =1050
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1305
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9780
                            Top =1050
                            Width =1260
                            Height =240
                            Name ="Label80"
                            Caption ="Trailer Width:"
                            LayoutCachedLeft =9780
                            LayoutCachedTop =1050
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =1290
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11100
                    Top =1350
                    Width =1500
                    Height =255
                    TabIndex =50
                    BorderColor =13017476
                    Name ="TrailerTaxAxle"
                    ControlSource ="TrailerTaxAxle"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11100
                    LayoutCachedTop =1350
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1605
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9780
                            Top =1350
                            Width =1260
                            Height =240
                            Name ="Label81"
                            Caption ="Trailer Tax Axle:"
                            LayoutCachedLeft =9780
                            LayoutCachedTop =1350
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =1590
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =12780
                    Top =1650
                    Width =1094
                    Height =255
                    TabIndex =51
                    BorderColor =13017476
                    Name ="TrailerMfgRebate"
                    ControlSource ="TrailerMfgRebate"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =12780
                    LayoutCachedTop =1650
                    LayoutCachedWidth =13874
                    LayoutCachedHeight =1905
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9780
                            Top =1665
                            Width =1260
                            Height =240
                            Name ="Label82"
                            Caption ="Mfg Rebate:"
                            LayoutCachedLeft =9780
                            LayoutCachedTop =1665
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =1905
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =4800
                    Top =4050
                    Width =1095
                    Height =240
                    Name ="Label83"
                    Caption ="Other1:"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =4050
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =4290
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =4800
                    Top =4335
                    Width =1095
                    Height =240
                    Name ="Label84"
                    Caption ="Other2:"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =4335
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =4575
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =4800
                    Top =4635
                    Width =1095
                    Height =240
                    Name ="Label85"
                    Caption ="Other3:"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =4635
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =4875
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    Left =12839
                    Top =450
                    Width =1035
                    Height =240
                    ForeColor =8421504
                    Name ="Label86"
                    Caption ="Rebate $"
                    LayoutCachedLeft =12839
                    LayoutCachedTop =450
                    LayoutCachedWidth =13874
                    LayoutCachedHeight =690
                End
                Begin Label
                    OverlapFlags =215
                    Left =9780
                    Top =390
                    Width =2580
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label87"
                    Caption ="Trailer Specs"
                    LayoutCachedLeft =9780
                    LayoutCachedTop =390
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =675
                End
                Begin Label
                    OverlapFlags =215
                    Left =60
                    Top =390
                    Width =2580
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Tractor Specs"
                    LayoutCachedLeft =60
                    LayoutCachedTop =390
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =675
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =87
                    Left =4680
                    Top =705
                    Width =15
                    Height =4560
                    Name ="Line25"
                    LayoutCachedLeft =4680
                    LayoutCachedTop =705
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =5265
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =3750
                    Width =2160
                    Height =255
                    TabIndex =39
                    BorderColor =13017476
                    Name ="SRFanClutch"
                    ControlSource ="SRFanClutch"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3750
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =4005
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =3450
                    Width =2160
                    Height =255
                    TabIndex =37
                    BorderColor =13017476
                    Name ="SRFilter"
                    ControlSource ="SRFilter"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3450
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =3705
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9630
                    Top =4035
                    Width =4275
                    Height =1161
                    TabIndex =56
                    Name ="SRNotes"
                    ControlSource ="SRNotes"
                    StatusBarText ="added at Frank's request 5-6-10"

                    LayoutCachedLeft =9630
                    LayoutCachedTop =4035
                    LayoutCachedWidth =13905
                    LayoutCachedHeight =5196
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9630
                            Top =3750
                            Width =600
                            Height =240
                            Name ="Label97"
                            Caption ="Notes:"
                            LayoutCachedLeft =9630
                            LayoutCachedTop =3750
                            LayoutCachedWidth =10230
                            LayoutCachedHeight =3990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11220
                    Top =2625
                    Height =255
                    TabIndex =52
                    BorderColor =13017476
                    Name ="SRSellingPrice"
                    ControlSource ="SRSellingPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="added at Frank's request 5-13-10"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11220
                    LayoutCachedTop =2625
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =2880
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9615
                            Top =2625
                            Width =1515
                            Height =240
                            Name ="Label98"
                            Caption ="Selling Price:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =2625
                            LayoutCachedWidth =11130
                            LayoutCachedHeight =2865
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11220
                    Top =2925
                    Height =255
                    TabIndex =53
                    BorderColor =13017476
                    Name ="SRListPrice"
                    ControlSource ="SRListPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="added at Frank's request 5-13-10"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11220
                    LayoutCachedTop =2925
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =3180
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9615
                            Top =2925
                            Width =1515
                            Height =240
                            Name ="Label99"
                            Caption ="List Price:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =2925
                            LayoutCachedWidth =11130
                            LayoutCachedHeight =3165
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11220
                    Top =3210
                    Height =255
                    TabIndex =54
                    BorderColor =13017476
                    Name ="SROtherPrice"
                    ControlSource ="SROtherPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="added at Frank's request 5-13-10"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11220
                    LayoutCachedTop =3210
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =3465
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9615
                            Top =3210
                            Width =1515
                            Height =240
                            Name ="Label100"
                            Caption ="Delivery/Tax/Other:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =3210
                            LayoutCachedWidth =11130
                            LayoutCachedHeight =3450
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11220
                    Top =3495
                    Height =255
                    TabIndex =55
                    BorderColor =13017476
                    Name ="SRDiscountPrice"
                    ControlSource ="SRDiscountPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="added 5-13-10 - formula: Selling Price (divided by) List Price (plus) Delivery P"
                        "rice/Tax/Other , then (subtract) 1% = DISCOUNT PRICE"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001010000 ,
                        0x22b14c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =11220
                    LayoutCachedTop =3495
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =3750
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000101000022b14c00ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9615
                            Top =3495
                            Width =1515
                            Height =240
                            Name ="Label101"
                            Caption ="Discount Price:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =3495
                            LayoutCachedWidth =11130
                            LayoutCachedHeight =3735
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1245
                    Top =1650
                    Height =255
                    TabIndex =5
                    BorderColor =13017476
                    Name ="SREngineModel"
                    ControlSource ="SREngineModel"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =1650
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =1905
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =60
                            Top =1650
                            Width =1080
                            Height =240
                            Name ="Label102"
                            Caption ="Engine Model:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1650
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1890
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1230
                    Top =4860
                    Width =2160
                    Height =255
                    TabIndex =26
                    BorderColor =13017476
                    Name ="SRRoofFairing"
                    ControlSource ="SRRoofFairing"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =1230
                    LayoutCachedTop =4860
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =5115
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =30
                            Top =4860
                            Width =1095
                            Height =240
                            Name ="Label103"
                            Caption ="Roof Fairing:"
                            LayoutCachedLeft =30
                            LayoutCachedTop =4860
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =5100
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6015
                    Top =1650
                    Width =2160
                    Height =255
                    TabIndex =30
                    BorderColor =13017476
                    Name ="SRSideExtenders"
                    ControlSource ="SRSideExtenders"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =1650
                    LayoutCachedWidth =8175
                    LayoutCachedHeight =1905
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4815
                            Top =1650
                            Width =1095
                            Height =240
                            Name ="Label104"
                            Caption ="Side Extend.:"
                            LayoutCachedLeft =4815
                            LayoutCachedTop =1650
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =1890
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6015
                    Top =750
                    Width =2160
                    Height =255
                    TabIndex =27
                    BorderColor =13017476
                    Name ="SRSideFarings"
                    ControlSource ="SRSideFarings"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =750
                    LayoutCachedWidth =8175
                    LayoutCachedHeight =1005
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4815
                            Top =750
                            Width =1095
                            Height =240
                            Name ="Label105"
                            Caption ="Side Farings:"
                            LayoutCachedLeft =4815
                            LayoutCachedTop =750
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =990
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6015
                    Top =1050
                    Width =2160
                    Height =255
                    TabIndex =28
                    BorderColor =13017476
                    Name ="SRSleeperSize"
                    ControlSource ="SRSleeperSize"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =1050
                    LayoutCachedWidth =8175
                    LayoutCachedHeight =1305
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4815
                            Top =1050
                            Width =1095
                            Height =240
                            Name ="Label106"
                            Caption ="Sleeper Size:"
                            LayoutCachedLeft =4815
                            LayoutCachedTop =1050
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =1290
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6015
                    Top =1350
                    Width =2160
                    Height =255
                    TabIndex =29
                    BorderColor =13017476
                    Name ="SRDoubleBunk"
                    ControlSource ="SRDoubleBunk"
                    StatusBarText ="added at Frank's request 5-5-10"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =1350
                    LayoutCachedWidth =8175
                    LayoutCachedHeight =1605
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4815
                            Top =1350
                            Width =1095
                            Height =240
                            Name ="Label107"
                            Caption ="Double Bunk:"
                            LayoutCachedLeft =4815
                            LayoutCachedTop =1350
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =1590
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4170
                    Top =30
                    Width =7935
                    Height =255
                    TabIndex =58
                    BorderColor =13017476
                    ForeColor =1279872587
                    Name ="SREqSpecLink"
                    ControlSource ="SREqSpecLink"
                    StatusBarText ="link to Equipment Spec"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4170
                    LayoutCachedTop =30
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =285
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2805
                            Top =45
                            Width =1290
                            Height =240
                            Name ="Label64"
                            Caption ="Equipment Spec:"
                            LayoutCachedLeft =2805
                            LayoutCachedTop =45
                            LayoutCachedWidth =4095
                            LayoutCachedHeight =285
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12165
                    Width =1425
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =59
                    ForeColor =-2147483615
                    Name ="AddDocBtn"
                    Caption ="Click to add file"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12165
                    LayoutCachedWidth =13590
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12225
                    Top =2310
                    Width =1710
                    Height =255
                    FontWeight =700
                    TabIndex =60
                    Name ="SRRebateTotal"
                    ControlSource ="SRRebateTotal"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12225
                    LayoutCachedTop =2310
                    LayoutCachedWidth =13935
                    LayoutCachedHeight =2565
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9600
                            Top =2325
                            Width =2580
                            Height =240
                            FontWeight =700
                            Name ="Label69"
                            Caption ="RebateTotal (tractor/trailer):"
                            LayoutCachedLeft =9600
                            LayoutCachedTop =2325
                            LayoutCachedWidth =12180
                            LayoutCachedHeight =2565
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1485
                    Top =60
                    Width =1125
                    TabIndex =57
                    BorderColor =13017476
                    Name ="SRSpecDate"
                    ControlSource ="SRSpecDate"
                    StatusBarText ="this is the date on the specification"

                    LayoutCachedLeft =1485
                    LayoutCachedTop =60
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =45
                            Top =75
                            Width =1425
                            Height =240
                            Name ="Label96"
                            Caption ="Specification Date:"
                            LayoutCachedLeft =45
                            LayoutCachedTop =75
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =315
                        End
                    End
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1845
            Name ="FormFooter"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =1755
                    Top =105
                    Width =7650
                    Height =1575
                    Name ="Quotes_sub"
                    SourceObject ="Form.Quotes_sub"
                    LinkChildFields ="specREF"
                    LinkMasterFields ="SRID"

                    LayoutCachedLeft =1755
                    LayoutCachedTop =105
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1680
                End
                Begin Label
                    OverlapFlags =85
                    Left =90
                    Top =90
                    Width =1605
                    Height =525
                    FontSize =10
                    FontWeight =700
                    Name ="Label94"
                    Caption ="Quotes for this Spec Review:"
                    LayoutCachedLeft =90
                    LayoutCachedTop =90
                    LayoutCachedWidth =1695
                    LayoutCachedHeight =615
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =765
                    Width =1305
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =32768
                    Name ="AddQuoteBtn"
                    Caption ="Add a Quote"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =120
                    LayoutCachedTop =765
                    LayoutCachedWidth =1425
                    LayoutCachedHeight =1155
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1245
                    Width =1290
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =32768
                    Name ="RefreshBtn"
                    Caption ="Refresh"
                    OnClick ="Refresh"
                    FontName ="Tahoma"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1245
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =1605
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10905
                    Top =885
                    Width =1770
                    Height =255
                    TabIndex =5
                    BorderColor =13017476
                    Name ="SRCPSentDate1"
                    ControlSource ="SRCPSentDate1"
                    Format ="Short Date"
                    StatusBarText ="comp sent date"

                    LayoutCachedLeft =10905
                    LayoutCachedTop =885
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9735
                            Top =885
                            Width =855
                            Height =240
                            Name ="Label113"
                            Caption ="Sent Spec:"
                            LayoutCachedLeft =9735
                            LayoutCachedTop =885
                            LayoutCachedWidth =10590
                            LayoutCachedHeight =1125
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10905
                    Top =585
                    Width =1770
                    Height =255
                    TabIndex =4
                    BorderColor =13017476
                    Name ="SRCPReqDate1"
                    ControlSource ="SRCPReqDate1"
                    Format ="Short Date"

                    LayoutCachedLeft =10905
                    LayoutCachedTop =585
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9735
                            Top =585
                            Width =1125
                            Height =240
                            Name ="Label114"
                            Caption ="Request Date:"
                            LayoutCachedLeft =9735
                            LayoutCachedTop =585
                            LayoutCachedWidth =10860
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10905
                    Top =1185
                    Width =1770
                    Height =255
                    TabIndex =6
                    BorderColor =13017476
                    Name ="SRCPRecdDate1"
                    ControlSource ="SRCPRecdDate1"
                    Format ="Short Date"
                    StatusBarText ="comp recd date"

                    LayoutCachedLeft =10905
                    LayoutCachedTop =1185
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9735
                            Top =1185
                            Width =1020
                            Height =240
                            Name ="Label115"
                            Caption ="Rec'd Quote:"
                            LayoutCachedLeft =9735
                            LayoutCachedTop =1185
                            LayoutCachedWidth =10755
                            LayoutCachedHeight =1425
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10890
                    Top =1505
                    Width =1770
                    Height =255
                    TabIndex =7
                    BorderColor =13017476
                    Name ="SRCPCost1"
                    ControlSource ="SRCPCost1"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="comp cost"

                    LayoutCachedLeft =10890
                    LayoutCachedTop =1505
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =1760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9750
                            Top =1490
                            Width =870
                            Height =240
                            Name ="Label116"
                            Caption ="Price:"
                            LayoutCachedLeft =9750
                            LayoutCachedTop =1490
                            LayoutCachedWidth =10620
                            LayoutCachedHeight =1730
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =9720
                    Width =2580
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label117"
                    Caption ="Comp. Equip. Pricing"
                    LayoutCachedLeft =9720
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =285
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10905
                    Top =315
                    Width =2595
                    FontSize =8
                    TabIndex =3
                    BorderColor =13017476
                    Name ="SRCompany1"
                    ControlSource ="SRCompany1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorName FROM Vendors ORDER BY Vendors.VendorName;"
                    StatusBarText ="comp pricing company"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =10905
                    LayoutCachedTop =315
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9735
                            Top =315
                            Width =900
                            Height =240
                            Name ="Label112"
                            Caption ="Vendor:"
                            ControlTipText ="Missing Vendor? Request Purchasing to set up"
                            LayoutCachedLeft =9735
                            LayoutCachedTop =315
                            LayoutCachedWidth =10635
                            LayoutCachedHeight =555
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Spec_Review_sub.cls"
