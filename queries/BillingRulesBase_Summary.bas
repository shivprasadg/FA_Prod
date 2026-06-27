Operation =1
Option =2
Where ="(((Bills.BillID)=1476) AND ((vw_SixKeys.clientID)=4971) AND ((vw_SixKeys.MLOrig)"
    "=\"FA\") AND ((Units.UnitStatus)=\"A\"))"
Begin InputTables
    Name ="Bills"
    Name ="vw_SixKeys"
    Name ="BillingRulesClient"
    Name ="Units"
    Name ="Locations"
    Name ="UnitOEC2"
    Name ="SchGrp"
    Name ="lookup_States"
    Name ="Schedule"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Bank_Assignee"
    Name ="BillingRules_FeeLabels"
End
Begin OutputColumns
    Expression ="Bills.BillID"
    Expression ="Bills.BillBank"
    Alias ="InvoiceNumber"
    Expression ="Bills.BillNo"
    Expression ="Bills.Billdt"
    Expression ="Bills.Billfrom"
    Expression ="Bills.Billto"
    Expression ="Bills.BillDueDt"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.UnitNum"
    Expression ="Units.UnitStatus"
    Expression ="vw_SixKeys.VIN"
    Alias ="Desc"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & \" \" & [UnitType]"
    Expression ="UnitOEC2.UnitCost"
    Expression ="UnitOEC2.TotalCost"
    Alias ="TaxExempt"
    Expression ="Units.UnitTaxExempt"
    Alias ="GaragedLocation"
    Expression ="[Locations].[LocationName] & \" - \" & [Locations].[LState]"
    Expression ="lookup_States.StTaxType"
    Expression ="Locations.LTaxRate"
    Alias ="UnitAccepted"
    Expression ="Units.UnitAcceptDate"
    Alias ="BLCD"
    Expression ="SchGrp.BLCD"
    Alias ="BillToAddress"
    Expression ="[Locations_1].[LocationStreet] & Chr(13) & Chr(10) & [locations_1].[LAddrLine1] "
        "& Chr(13) & Chr(10) & IIf(IsNull([locations_1].[LAddrLine2]),Null,[locations_1]."
        "[LAddrLine2] & Chr(13) & Chr(10)) & [locations_1].[LCity] & \" \" & [locations_1"
        "].[LState] & \" \" & [locations_1].[LZip]"
    Expression ="Bank_Assignee.BankAddress"
    Alias ="FARemit"
    Expression ="\"Fleet Advantage, LLC\" & Chr(13) & Chr(10) & \"401 East Las Olas Blvd. Suite 1"
        "720\" & Chr(13) & Chr(10) & \"Fort Lauderdale, FL 33301\""
    Alias ="RemitAddress"
    Expression ="IIf([RemitTo]=2,[FARemit],[BankAddress])"
    Alias ="Rent"
    Expression ="CCur(IIf([BillingRulesClient].[BillType]>5,[UnitsRent],[unitsrent]))"
    Alias ="RentTax"
    Expression ="IIf([UnitTaxExempt]=True,0,[locations].[LTaxRate]*[Rent])"
    Alias ="FeeTax"
    Expression ="IIf([FeeTaxable]=0 Or [UnitTaxExempt]=True,0,([Feevalue]*[locations].[LTaxRate])"
        ")"
    Alias ="TaxAmt"
    Expression ="[RentTax]+[FeeTax]"
    Alias ="BankContract"
    Expression ="IIf([ShowContract]=1,[SchBnkClNO] & \"-\" & [SgrpBnkNo],\"hide\")"
    Expression ="Bills.DueonRec"
    Alias ="FeeEndDate"
    Expression ="IIf([feeterm]=0,DateAdd(\"m\",[vw_SixKeys].[LeaseTermPrimary],[BLCD]),DateAdd(\""
        "m\",[FeeTerm],[BLCD]))-1"
    Alias ="Fee"
    Expression ="CCur(IIf([FeeToUse]=0 Or [FeeEndDate]<[Billfrom],0,[FeeValue]))"
    Alias ="LineTotal"
    Expression ="CCur([Rent]+[TaxAmt]+[Fee])"
    Alias ="FeeLabel"
    Expression ="BillingRules_FeeLabels.RuleString"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="BillingRulesClient"
    Expression ="vw_SixKeys.clientID = BillingRulesClient.ClientId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="BillingRulesClient"
    Expression ="vw_SixKeys.SchID = BillingRulesClient.ScheduleId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Bills"
    RightTable ="vw_SixKeys"
    Expression ="Bills.BillSchRef = vw_SixKeys.SchID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="UnitOEC2"
    Expression ="vw_SixKeys.AssetID = UnitOEC2.AssetId"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Schedule"
    Expression ="vw_SixKeys.SchID = Schedule.SchID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="Locations_1"
    Expression ="Schedule.SchBillAdd = Locations_1.LocationID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Bank_Assignee"
    Expression ="SchGrp.Assignee_Bank = Bank_Assignee.ShortName"
    Flag =1
    LeftTable ="BillingRulesClient"
    RightTable ="BillingRules_FeeLabels"
    Expression ="BillingRulesClient.FeeLabel = BillingRules_FeeLabels.RuleValue"
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
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3060"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitOEC2.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitAccepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaragedLocation"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank_Assignee.BankAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankContract"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARemit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemitAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeLabel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =26
    Right =1420
    Bottom =582
    Left =-1
    Top =-1
    Right =1344
    Bottom =255
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =4
        Top =21
        Right =148
        Bottom =165
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =155
        Top =19
        Right =331
        Bottom =452
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =360
        Top =157
        Right =504
        Bottom =301
        Top =0
        Name ="BillingRulesClient"
        Name =""
    End
    Begin
        Left =362
        Top =309
        Right =506
        Bottom =453
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =526
        Top =381
        Right =670
        Bottom =525
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =361
        Top =458
        Right =505
        Bottom =657
        Top =0
        Name ="UnitOEC2"
        Name =""
    End
    Begin
        Left =515
        Top =154
        Right =659
        Bottom =298
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =689
        Top =388
        Right =833
        Bottom =532
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =514
        Top =6
        Right =658
        Bottom =150
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =688
        Top =7
        Right =832
        Bottom =151
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =679
        Top =168
        Right =823
        Bottom =312
        Top =0
        Name ="Bank_Assignee"
        Name =""
    End
    Begin
        Left =1000
        Top =147
        Right =1144
        Bottom =291
        Top =0
        Name ="BillingRules_FeeLabels"
        Name =""
    End
End
