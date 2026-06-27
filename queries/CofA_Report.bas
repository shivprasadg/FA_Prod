Operation =1
Option =0
Begin InputTables
    Name ="CofA_Report_A"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="lookup_States"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="Locations"
    Alias ="Locations_3"
    Name ="lookup_groupTypes"
    Name ="CLASS"
    Name ="Contacts"
    Name ="MLAmends"
    Name ="ScheduleAmends"
End
Begin OutputColumns
    Expression ="CofA_Report_A.*"
    Alias ="TSt"
    Expression ="Locations.LState"
    Alias ="DSub"
    Expression ="Locations_1.LocSubLessee"
    Alias ="Dst1"
    Expression ="Locations_1.LocationStreet"
    Alias ="DST2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="DCITY"
    Expression ="Locations_1.LCity"
    Alias ="DCNTY"
    Expression ="Locations_1.LCounty"
    Alias ="DST"
    Expression ="Locations_1.LState"
    Alias ="DzIP"
    Expression ="Locations_1.LZip"
    Expression ="lookup_States.StTaxType"
    Alias ="TaxType"
    Expression ="IIf([UnitTaxExempt]=Yes,\"Exempt\",[StTaxType])"
    Alias ="RSub"
    Expression ="Locations_2.LocSubLessee"
    Alias ="Rst1"
    Expression ="Locations_2.LocationStreet"
    Alias ="Rst2"
    Expression ="Locations_2.LAddrLine1"
    Alias ="RCity"
    Expression ="Locations_2.LCity"
    Alias ="RST"
    Expression ="Locations_2.LState"
    Alias ="RZIP"
    Expression ="Locations_2.LZip"
    Alias ="DelSub"
    Expression ="Locations_3.LocSubLessee"
    Alias ="DelSt1"
    Expression ="Locations_3.LocationStreet"
    Alias ="DelSt2"
    Expression ="Locations_3.LAddrLine1"
    Alias ="DelCity"
    Expression ="Locations_3.LCity"
    Alias ="DelSt"
    Expression ="Locations_3.LState"
    Alias ="DelZip"
    Expression ="Locations_3.LZip"
    Expression ="CLASS.TITLED"
    Expression ="MLAmends.MLAmtNO"
    Expression ="MLAmends.MLAmdDt"
    Expression ="ScheduleAmends.AmendmentNO"
    Expression ="ScheduleAmends.AmendmentDate"
    Expression ="Contacts.contactTitle"
    Alias ="FullNm"
    Expression ="[contactFirst] & \" \" & [contactMiddle] & IIf(IsNull([contactMiddle]),Null,\".\""
        ") & \" \" & [Contactlast] & \" \" & [contactsuffix]"
    Alias ="MLRMo"
    Expression ="IIf(IsNull([MLAmdDt]),\"N/A\",MonthName(DatePart(\"m\",[MLAmdDt])))"
    Alias ="MLRday"
    Expression ="IIf(IsNull([MLAmdDt]),\"N/A\",MonthName(DatePart(\"d\",[MLAmdDt])))"
    Alias ="MLRYr"
    Expression ="IIf(IsNull([MLAmdDt]),\"N/A\",MonthName(DatePart(\"yyyy\",[MLAmdDt])))"
    Alias ="Shmo"
    Expression ="IIf(IsNull([scheddt]),\"N/A\",MonthName(DatePart(\"mm\",[scheddt])))"
    Alias ="Schday"
    Expression ="IIf(IsNull([scheddt]),\"N/A\",MonthName(DatePart(\"dd\",[scheddt])))"
    Alias ="SchYr"
    Expression ="IIf(IsNull([scheddt]),\"N/A\",MonthName(DatePart(\"yyyy\",[scheddt])))"
    Alias ="GarLOC"
    Expression ="Locations_1.LTaxRate"
    Expression ="Locations_1.LCountry"
    Expression ="Locations_1.LTaxRate"
End
Begin Joins
    LeftTable ="CofA_Report_A"
    RightTable ="Locations"
    Expression ="CofA_Report_A.UnitTitleLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="Locations_1"
    Expression ="CofA_Report_A.UnitGarageLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Locations_1"
    RightTable ="lookup_States"
    Expression ="Locations_1.LState = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="Locations_2"
    Expression ="CofA_Report_A.UnitRegLocationRef = Locations_2.LocationID"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="Locations_3"
    Expression ="CofA_Report_A.UnitDeliveryLocationRef = Locations_3.LocationID"
    Flag =2
    LeftTable ="lookup_groupTypes"
    RightTable ="CLASS"
    Expression ="lookup_groupTypes.grpClass = CLASS.CLASS"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="lookup_groupTypes"
    Expression ="CofA_Report_A.UnitType = lookup_groupTypes.groupType"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="Contacts"
    Expression ="CofA_Report_A.CofASigner = Contacts.contactID"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="MLAmends"
    Expression ="CofA_Report_A.MLID = MLAmends.MLRef"
    Flag =2
    LeftTable ="CofA_Report_A"
    RightTable ="ScheduleAmends"
    Expression ="CofA_Report_A.SchID = ScheduleAmends.ScheduleID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CofA_Report_A.Units.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.GrpCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.MLYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.MLMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.mlday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.PRINTCOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Mileage.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dst1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCNTY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DzIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rst1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RZIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelSt1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CLASS.TITLED"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLRMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLRday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLRYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shmo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Schedule.SchedDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GarLOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UNITTAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Units.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_A.Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1268
    Bottom =825
    Left =-1
    Top =-1
    Right =1236
    Bottom =477
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =23
        Top =9
        Right =222
        Bottom =517
        Top =0
        Name ="CofA_Report_A"
        Name =""
    End
    Begin
        Left =731
        Top =7
        Right =875
        Bottom =151
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =886
        Top =30
        Right =1032
        Bottom =320
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =1063
        Top =284
        Right =1207
        Bottom =428
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =1051
        Top =51
        Right =1216
        Bottom =197
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =1234
        Top =83
        Right =1378
        Bottom =227
        Top =0
        Name ="Locations_3"
        Name =""
    End
    Begin
        Left =631
        Top =156
        Right =775
        Bottom =300
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
    Begin
        Left =808
        Top =334
        Right =954
        Bottom =456
        Top =0
        Name ="CLASS"
        Name =""
    End
    Begin
        Left =632
        Top =313
        Right =776
        Bottom =457
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =438
        Top =291
        Right =582
        Bottom =435
        Top =0
        Name ="MLAmends"
        Name =""
    End
    Begin
        Left =270
        Top =12
        Right =414
        Bottom =156
        Top =0
        Name ="ScheduleAmends"
        Name =""
    End
End
