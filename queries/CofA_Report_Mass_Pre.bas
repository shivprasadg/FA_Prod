Operation =1
Option =0
Begin InputTables
    Name ="Bulk_Update_WorkTable"
    Name ="CofA_Report_X"
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
    Name ="ScheduleAmends"
    Name ="ContactsIncumbent"
End
Begin OutputColumns
    Expression ="CofA_Report_X.*"
    Alias ="TitleLocName"
    Expression ="Locations.LocationName"
    Alias ="TitleLoc"
    Expression ="[Locations].[LocationStreet] & IIf(IsNull([Locations].[LocSubLessee]),Null,Chr(1"
        "3) & Chr(10) & [Locations].[LocSubLessee]) & Chr(13) & Chr(10) & [Locations].[LA"
        "ddrLine1] & Chr(13) & Chr(10) & [Locations].[LCity] & \"  \" & [Locations].[LSta"
        "te] & \" \" & IIf(Len([Locations].[LZip])>5,Left([Locations].[LZip],5) & \"-\" &"
        " Right([Locations].[LZip],4),[Locations].[LZip])"
    Alias ="GarageLoc"
    Expression ="[Locations_1].[LocationStreet] & IIf(IsNull([Locations_1].[LocSubLessee]),Null,C"
        "hr(13) & Chr(10) & [Locations_1].[LocSubLessee]) & Chr(13) & Chr(10) & [Location"
        "s_1].[LAddrLine1] & Chr(13) & Chr(10) & [Locations_1].[LCity] & \"  \" & [Locati"
        "ons_1].[LState] & \" \" & IIf(Len([Locations_1].[LZip])>5,Left([Locations_1].[LZ"
        "ip],5) & \"-\" & Right([Locations_1].[LZip],4),[Locations_1].[LZip])"
    Alias ="RegLoc"
    Expression ="[Locations_2].[LocationStreet] & IIf(IsNull([Locations_2].[LocSubLessee]),Null,C"
        "hr(13) & Chr(10) & [Locations_2].[LocSubLessee]) & Chr(13) & Chr(10) & [Location"
        "s_2].[LAddrLine1] & Chr(13) & Chr(10) & [Locations_2].[LCity] & \"  \" & [Locati"
        "ons_2].[LState] & \" \" & IIf(Len([Locations_2].[LZip])>5,Left([Locations_2].[LZ"
        "ip],5) & \"-\" & Right([Locations_2].[LZip],4),[Locations_2].[LZip])"
    Alias ="DeliveryLoc"
    Expression ="[Locations_3].[LocationStreet] & IIf(IsNull([Locations_3].[LocSubLessee]),Null,C"
        "hr(13) & Chr(10) & [Locations_3].[LocSubLessee]) & Chr(13) & Chr(10) & [Location"
        "s_3].[LAddrLine1] & Chr(13) & Chr(10) & [Locations_3].[LCity] & \"  \" & [Locati"
        "ons_3].[LState] & \" \" & IIf(Len([Locations_3].[LZip])>5,Left([Locations_3].[LZ"
        "ip],5) & \"-\" & Right([Locations_3].[LZip],4),[Locations_3].[LZip])"
    Alias ="TitledState"
    Expression ="Locations.LState"
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
    Expression ="IIf([Locations_1].[LCounty] Like \"*county*\",SplitString([Locations_1].[LCounty"
        "],\" County\",0),[Locations_1].[LCounty])"
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
    Expression ="ScheduleAmends.AmendmentNO"
    Expression ="ScheduleAmends.AmendmentDate"
    Expression ="Contacts.contactTitle"
    Alias ="FullNm"
    Expression ="[contactFirst] & \" \" & [contactMiddle] & IIf(IsNull([contactMiddle]),Null,\".\""
        ") & \" \" & [Contactlast] & \" \" & [contactsuffix]"
    Alias ="MLRMo"
    Expression ="\"n/a\""
    Alias ="MLRday"
    Expression ="\"n/a\""
    Alias ="MLRYr"
    Expression ="\"n/a\""
    Alias ="Shmo"
    Expression ="\"n/a\""
    Alias ="Schday"
    Expression ="\"n/a\""
    Alias ="SchYr"
    Expression ="\"n/a\""
    Alias ="GarLOC"
    Expression ="Locations_1.LTaxRate"
    Expression ="Locations_1.LCountry"
    Expression ="Locations_1.LTaxRate"
    Expression ="ContactsIncumbent.ClientSignTitle"
    Expression ="ContactsIncumbent.ClientSignFullName"
    Alias ="UnitDesc"
    Expression ="[Desc] & \"   \" & IIf([sgrpplb]=False,Null,\"  Mi/Hr \") & IIf([sgrpplb]=False,"
        "Null,[mileage]) & IIf([sgrpplb]=False,Null,\" / \") & IIf([sgrpplb]=False,Null,["
        "hours]) & \"   \" & [UnitVIN]"
    Alias ="RegState"
    Expression ="DLookUp(\"StateLong\",\"Lookup_States\",\"StateAbbrev='\" & IIf([Locations_2].[L"
        "ocationName] Like \"IRP Reg*\",Trim(SplitString([Locations_2].[LocationName],\"I"
        "RP Reg -\",1)),[Locations_2].[LState]) & \"'\")"
    Alias ="ProjEOL"
    Expression ="DateAdd(\"m\",[LeaseTermPrimary],[BLCD]-1)"
End
Begin Joins
    LeftTable ="CofA_Report_X"
    RightTable ="Locations"
    Expression ="CofA_Report_X.UnitTitleLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="Locations_1"
    Expression ="CofA_Report_X.UnitGarageLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Locations_1"
    RightTable ="lookup_States"
    Expression ="Locations_1.LState = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="Locations_2"
    Expression ="CofA_Report_X.UnitRegLocationRef = Locations_2.LocationID"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="Locations_3"
    Expression ="CofA_Report_X.UnitDeliveryLocationRef = Locations_3.LocationID"
    Flag =2
    LeftTable ="lookup_groupTypes"
    RightTable ="CLASS"
    Expression ="lookup_groupTypes.grpClass = CLASS.CLASS"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="lookup_groupTypes"
    Expression ="CofA_Report_X.UnitType = lookup_groupTypes.groupType"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="Contacts"
    Expression ="CofA_Report_X.CofASigner = Contacts.contactID"
    Flag =2
    LeftTable ="CofA_Report_X"
    RightTable ="ScheduleAmends"
    Expression ="CofA_Report_X.SchID = ScheduleAmends.ScheduleID"
    Flag =2
    LeftTable ="Bulk_Update_WorkTable"
    RightTable ="CofA_Report_X"
    Expression ="Bulk_Update_WorkTable.UnitID = CofA_Report_X.UnitID"
    Flag =1
    LeftTable ="CofA_Report_X"
    RightTable ="ContactsIncumbent"
    Expression ="CofA_Report_X.CofASigner = ContactsIncumbent.contactID"
    Flag =2
End
Begin OrderBy
    Expression ="CofA_Report_X.vin4"
    Flag =0
    Expression ="Locations_3.LAddrLine1"
    Flag =0
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
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCNTY"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3945"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DelSt2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Schday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GarLOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MstrLease.MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Schedule.SchedDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="CofA_Report_X.GrpCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientDBA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientZip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MLYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MLMo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="CofA_Report_X.mlday"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UNITTAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactsIncumbent.ClientSignTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactsIncumbent.ClientSignFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="CofA_Report_X.CorpState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.[StTaxType]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="8835"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitledState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveryLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GarageLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.UnitLastMileageRead.ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.CofASigner"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="CofA_Report_X.ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.MakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.RENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.Schedule.SubLeasee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.UnitLastMileageRead.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.UnitLastMileageRead.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleAmends.AmendmentNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ScheduleAmends.AmendmentDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CofA_Report_X.PLBMilesHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.plblabel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.PLBMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.PLBHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleLocName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_X.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjEOL"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =216
    Top =30
    Right =1902
    Bottom =815
    Left =-1
    Top =-1
    Right =1668
    Bottom =485
    Left =191
    Top =0
    ColumnsShown =539
    Begin
        Left =-193
        Top =0
        Right =-49
        Bottom =144
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
    Begin
        Left =74
        Top =-7
        Right =270
        Bottom =343
        Top =0
        Name ="CofA_Report_X"
        Name =""
    End
    Begin
        Left =540
        Top =7
        Right =684
        Bottom =151
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =427
        Top =-3
        Right =573
        Bottom =192
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =974
        Top =247
        Right =1118
        Bottom =391
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =704
        Top =223
        Right =869
        Bottom =369
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =547
        Top =188
        Right =691
        Bottom =332
        Top =0
        Name ="Locations_3"
        Name =""
    End
    Begin
        Left =180
        Top =343
        Right =324
        Bottom =487
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
    Begin
        Left =343
        Top =345
        Right =489
        Bottom =467
        Top =0
        Name ="CLASS"
        Name =""
    End
    Begin
        Left =212
        Top =270
        Right =356
        Bottom =414
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =1083
        Top =12
        Right =1227
        Bottom =156
        Top =0
        Name ="ScheduleAmends"
        Name =""
    End
    Begin
        Left =367
        Top =168
        Right =521
        Bottom =283
        Top =0
        Name ="ContactsIncumbent"
        Name =""
    End
End
