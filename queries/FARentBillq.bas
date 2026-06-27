Operation =1
Option =0
Having ="(((Bills.BillID)=Forms!Billing!BillID) And ((SchGrp.BLCD) Is Not Null) And ((Sch"
    "Grp.Assignee_Bank) Is Null) And ((SchGrp.LoanBank) Is Null)) Or (((Bills.BillID)"
    "=Forms!Billing!BillID) And ((SchGrp.BLCD) Is Not Null) And ((SchGrp.Assignee_Ban"
    "k) Is Not Null)) Or (((Bills.BillID)=Forms!Billing!BillID) And ((SchGrp.BLCD) Is"
    " Not Null) And ((SchGrp.Assignee_Bank) Is Null) And ((SchGrp.LoanBank) Is Not Nu"
    "ll) And ((SchGrp.SGrpFABill)=True))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="PDstdtq"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Bills"
    Name ="SchGrp"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
    Name ="ChildPartSummaryQuery"
    Name ="BillAdjT"
    Name ="LocationsRemit"
    Name ="Bank_Assignee"
End
Begin OutputColumns
    Expression ="Bills.BillID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationName"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [Un"
        "itSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [UnitType])"
    Alias ="UCost"
    Expression ="[UnitFinalCost]+IIf(IsNull([sumofchildfinalcost]),0,[sumofchildfinalcost])"
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.UnitAcceptDate"
    Expression ="SchGrp.BLCD_Projected"
    Alias ="BLoST3"
    Expression ="Locations_1.LAddrLine2"
    Alias ="BLocStreet"
    Expression ="Locations_1.LocationStreet"
    Alias ="BLoST2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="BLocCity"
    Expression ="Locations_1.LCity"
    Alias ="BLocSt"
    Expression ="Locations_1.LState"
    Alias ="BLocZip"
    Expression ="Locations_1.LZip"
    Expression ="Clients.clientCompanyName"
    Expression ="lookup_States.StTaxType"
    Expression ="Bills.BillNo"
    Expression ="Bills.Billdt"
    Expression ="Bills.Billfrom"
    Expression ="Bills.Billto"
    Expression ="Bills.BillDueDt"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Alias ="FSLFee"
    Expression ="IIf([UNITSRENT]>0,IIf([ScheduleGroupPaymentTerm]>0,IIf([Billto]<DateAdd(\"m\",[S"
        "cheduleGroupPaymentTerm],[BLCD]),[SchGrp].[SGrpFSLPmt],0),[SchGrp].[SGrpFSLPmt])"
        ",0)"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.UNITSRENT"
    Expression ="Units.UnitTax"
    Alias ="UnitsBillRent"
    Expression ="IIf([unitstatus]=\"A\" Or IsNull([unitoffleasedt]) Or [unitoffleasedt]<[billfrom"
        "],[unitrent],0)"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Alias ="TaxAmt"
    Expression ="IIf([unittaxexempt]=-1,0,Round([locations].[LTaxRate]*[unitsrent],2))"
    Expression ="Locations.LTaxRate"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.LoanBank"
    Expression ="SchGrp.SGrpFABill"
    Expression ="Units.unitoffleasedt"
    Alias ="Bill1"
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    Alias ="Bill2"
    Expression ="UCase([invfld2desc] & \" \" & [invfld2] & \" \" & [lesseeacct])"
    Alias ="BLocSt2"
    Expression ="Locations_1.LAddrLine1"
    Expression ="Schedule.SchBnkClNO"
    Expression ="Clients.clientID"
    Expression ="SchGrp.SgrpBnkNo"
    Expression ="Locations_1.LAddrLine1"
    Expression ="Locations_1.LAddrLine2"
    Expression ="Locations_1.LAddrLine3"
    Alias ="SumOfBillAdj"
    Expression ="Sum(BillAdjT.BillAdj)"
    Alias ="DateFeesEnd"
    Expression ="DateAdd(\"m\",[ScheduleGroupPaymentTerm],[BLCD])"
    Alias ="ChargeFee"
    Expression ="IIf([Billto]<[DateFeesEnd],\"Yes\",\"No\")"
    Alias ="LineTot"
    Expression ="IIf(([unitstatus]=\"O\" Or [unitstatus]=\"S\" Or [unitstatus]=\"R\") And [unitof"
        "fleasedt]<[billfrom],0,[unitsrent])"
    Expression ="Bank_Assignee.ShortName"
    Alias ="Remitline1"
    Expression ="Bank_Assignee.BankName"
    Alias ="Remitline2"
    Expression ="Bank_Assignee.LAddrLine1"
    Alias ="Remitline3"
    Expression ="Bank_Assignee.LAddrLine2"
    Alias ="Remitline4"
    Expression ="[Bank_Assignee].[LCity] & \" \" & [Bank_Assignee].[LState] & \"  \" & Left([Bank"
        "_Assignee].[LZip],5)"
    Alias ="UnitCount"
    Expression ="IIf([UNITSRENT]>0,1,0)"
    Expression ="SchGrp.sgrpid"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="PDstdtq"
    Expression ="Schedule.SchID = PDstdtq.SchID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="Locations_1"
    Expression ="Schedule.SchBillAdd = Locations_1.LocationID"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="Bills"
    Expression ="Schedule.SchID = Bills.BillSchRef"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Units"
    Expression ="Locations.LocationID = Units.UnitGarageLocationRef"
    Flag =1
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildPartSummaryQuery"
    Expression ="Units.UnitID = ChildPartSummaryQuery.UnitREF"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Bills"
    RightTable ="BillAdjT"
    Expression ="Bills.BillID = BillAdjT.BillAdjBillref"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="LocationsRemit"
    Expression ="MstrLease.Client = LocationsRemit.ClientREF"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Bank_Assignee"
    Expression ="SchGrp.Assignee_Bank = Bank_Assignee.ShortName"
    Flag =2
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
End
Begin Groups
    Expression ="Bills.BillID"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Units.unitunitnum"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [Un"
        "itSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [UnitType])"
    GroupLevel =0
    Expression ="[UnitFinalCost]+IIf(IsNull([sumofchildfinalcost]),0,[sumofchildfinalcost])"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LCounty"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="SchGrp.BLCD_Projected"
    GroupLevel =0
    Expression ="Locations_1.LocationStreet"
    GroupLevel =0
    Expression ="Locations_1.LCity"
    GroupLevel =0
    Expression ="Locations_1.LState"
    GroupLevel =0
    Expression ="Locations_1.LZip"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="lookup_States.StTaxType"
    GroupLevel =0
    Expression ="Bills.BillNo"
    GroupLevel =0
    Expression ="Bills.Billdt"
    GroupLevel =0
    Expression ="Bills.Billfrom"
    GroupLevel =0
    Expression ="Bills.Billto"
    GroupLevel =0
    Expression ="Bills.BillDueDt"
    GroupLevel =0
    Expression ="Right([unitvin],6)"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="Units.UNITSRENT"
    GroupLevel =0
    Expression ="Units.UnitTax"
    GroupLevel =0
    Expression ="IIf([unitstatus]=\"A\" Or IsNull([unitoffleasedt]) Or [unitoffleasedt]<[billfrom"
        "],[unitrent],0)"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtFreq"
    GroupLevel =0
    Expression ="SchGrp.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="IIf([unittaxexempt]=-1,0,Round([locations].[LTaxRate]*[unitsrent],2))"
    GroupLevel =0
    Expression ="Locations.LTaxRate"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
    GroupLevel =0
    Expression ="SchGrp.LoanBank"
    GroupLevel =0
    Expression ="SchGrp.SGrpFABill"
    GroupLevel =0
    Expression ="Units.unitoffleasedt"
    GroupLevel =0
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    GroupLevel =0
    Expression ="UCase([invfld2desc] & \" \" & [invfld2] & \" \" & [lesseeacct])"
    GroupLevel =0
    Expression ="Schedule.SchBnkClNO"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="SchGrp.SgrpBnkNo"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine2"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine3"
    GroupLevel =0
    Expression ="DateAdd(\"m\",[ScheduleGroupPaymentTerm],[BLCD])"
    GroupLevel =0
    Expression ="IIf(([unitstatus]=\"O\" Or [unitstatus]=\"S\" Or [unitstatus]=\"R\") And [unitof"
        "fleasedt]<[billfrom],0,[unitsrent])"
    GroupLevel =0
    Expression ="Bank_Assignee.ShortName"
    GroupLevel =0
    Expression ="Bank_Assignee.BankName"
    GroupLevel =0
    Expression ="Bank_Assignee.LAddrLine1"
    GroupLevel =0
    Expression ="Bank_Assignee.LAddrLine2"
    GroupLevel =0
    Expression ="[Bank_Assignee].[LCity] & \" \" & [Bank_Assignee].[LState] & \"  \" & Left([Bank"
        "_Assignee].[LZip],5)"
    GroupLevel =0
    Expression ="IIf([UNITSRENT]>0,1,0)"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitType] & \" \""
        " & [UnitSubType]"
    GroupLevel =0
    Expression ="IIf([UNITSRENT]>0,IIf([ScheduleGroupPaymentTerm]>0,IIf([Billto]<DateAdd(\"m\",[S"
        "cheduleGroupPaymentTerm],SchGrp.BLCD),[SchGrp].[SGrpFSLPmt],0),[SchGrp].[SGrpFSL"
        "Pmt]),0)"
    GroupLevel =0
    Expression ="Clients.FAVenorNo"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine2"
    GroupLevel =0
    Expression ="SchGrp.sgrpid"
    GroupLevel =0
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
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTaxRt"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocCity"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocStreet"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxAmt"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billfrom"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UCost"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineTot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfBillAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank_Assignee.ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remitline1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remitline2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remitline3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remitline4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DateFeesEnd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ChargeFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="unitsbillrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitsrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =16
    Top =83
    Right =1165
    Bottom =771
    Left =-1
    Top =-1
    Right =1117
    Bottom =106
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="PDstdtq"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =624
        Top =156
        Right =768
        Bottom =300
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =816
        Top =156
        Right =960
        Bottom =300
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =48
        Top =300
        Right =192
        Bottom =444
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
    Begin
        Left =240
        Top =300
        Right =384
        Bottom =444
        Top =0
        Name ="BillAdjT"
        Name =""
    End
    Begin
        Left =432
        Top =300
        Right =576
        Bottom =444
        Top =0
        Name ="LocationsRemit"
        Name =""
    End
    Begin
        Left =624
        Top =300
        Right =768
        Bottom =444
        Top =0
        Name ="Bank_Assignee"
        Name =""
    End
End
