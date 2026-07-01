dbMemo "SQL" ="SELECT DISTINCT Bills.BillID, Bills.BillBank, Bills.BillNo AS InvoiceNumber, Bil"
    "ls.Billdt, Bills.Billfrom, Bills.Billto, Bills.BillDueDt, vw_SixKeys.ClientGroup"
    "Id, vw_SixKeys.Schedule, vw_SixKeys.UnitGroup, vw_SixKeys.clientID, vw_SixKeys.C"
    "ompanyName, vw_SixKeys.MLOrig, vw_SixKeys.MLNo, vw_SixKeys.AssetID, vw_SixKeys.U"
    "nitNum, Units.UnitStatus, vw_SixKeys.VIN, [unityr] & \" \" & [unitmake] & \" \" "
    "& [unitmodel] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [UnitSubType]) & \""
    " \" & [UnitType] AS [Desc], UnitOEC2.UnitCost, UnitOEC2.TotalCost, Units.UnitTax"
    "Exempt AS TaxExempt, [Locations].[LocationName] & \" - \" & [Locations].[LState]"
    " AS GaragedLocation, lookup_States.StTaxType, Locations.LTaxRate, Units.UnitAcce"
    "ptDate AS UnitAccepted, SchGrp.BLCD AS BLCD, [Locations_1].[LocationStreet] & Ch"
    "r(13) & Chr(10) & [locations_1].[LAddrLine1] & Chr(13) & Chr(10) & IIf(IsNull([l"
    "ocations_1].[LAddrLine2]),Null,[locations_1].[LAddrLine2] & Chr(13) & Chr(10)) &"
    " [locations_1].[LCity] & \" \" & [locations_1].[LState] & \" \" & [locations_1]."
    "[LZip] AS BillToAddress, Bank_Assignee.BankAddress, \"Fleet Advantage, LLC\" & C"
    "hr(13) & Chr(10) & \"401 East Las Olas Blvd. Suite 1720\" & Chr(13) & Chr(10) & "
    "\"Fort Lauderdale, FL 33301\" AS FARemit, BillingRulesClient.RemitTo, IIf([Remit"
    "To]=2,[FARemit],[BankAddress]) AS RemitAddress, CCur(IIf([BillingRulesClient].[B"
    "illType]>5,[UnitRent],[UnitsRent])) AS Rent, IIf([UnitTaxExempt]=True,0,[locatio"
    "ns].[LTaxRate]*[Rent]) AS RentTax, IIf([FeeTaxable]=0 Or [UnitTaxExempt]=True,0,"
    "([Feevalue]*[locations].[LTaxRate])) AS FeeTax, [RentTax]+[FeeTax] AS TaxAmt, II"
    "f([ShowContract]=1,[SchBnkClNO] & \"-\" & [SgrpBnkNo],\"hide\") AS BankContract,"
    " Bills.DueonRec, IIf([feeterm]=0,DateAdd(\"m\",vw_SixKeys.[LeaseTermPrimary],[BL"
    "CD]),DateAdd(\"m\",[FeeTerm],[BLCD]))-1 AS FeeEndDate, CCur(IIf([FeeToUse]=0 Or "
    "[FeeEndDate]<[Billfrom],0,[FeeValue])) AS Fee, CCur([Rent]+[TaxAmt]+[Fee]) AS Li"
    "neTotal, BillingRules_FeeLabels.RuleString AS FeeLabel, 1 AS UnitCount\015\012FR"
    "OM (((((((Bills INNER JOIN (((vw_SixKeys LEFT JOIN BillingRulesClient ON (vw_Six"
    "Keys.SchID = BillingRulesClient.ScheduleId) AND (vw_SixKeys.clientID = BillingRu"
    "lesClient.ClientId)) INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID) INNE"
    "R JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID) ON Bills"
    ".BillSchRef = vw_SixKeys.SchID) INNER JOIN UnitOEC2 ON vw_SixKeys.AssetID = Unit"
    "OEC2.AssetId) INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID) INNER JOIN"
    " lookup_States ON Locations.LState = lookup_States.StateAbbrev) INNER JOIN Sched"
    "ule ON vw_SixKeys.SchID = Schedule.SchID) INNER JOIN Locations AS Locations_1 ON"
    " Schedule.SchBillAdd = Locations_1.LocationID) INNER JOIN Bank_Assignee ON SchGr"
    "p.Assignee_Bank = Bank_Assignee.ShortName) LEFT JOIN BillingRules_FeeLabels ON B"
    "illingRulesClient.FeeLabel = BillingRules_FeeLabels.RuleValue\015\012WHERE (((Bi"
    "lls.BillID)=1476) AND ((vw_SixKeys.clientID)=4971) AND ((vw_SixKeys.MLOrig)=\"FA"
    "\") AND ((Units.UnitStatus)=\"A\"));\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2295"
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
        dbText "Name" ="Locations.LocTaxRt"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
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
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingRulesClient.RemitTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
End
