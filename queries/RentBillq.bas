Operation =1
Option =0
Where ="(((Bills.BillID)=Forms!Billing!BillingSubForm.Form!BillID) And ((SchGrp.BLCD) Is"
    " Not Null) And ((SchGrp.Assignee_Bank) Is Null Or (SchGrp.Assignee_Bank)=\"FA ll"
    "c\") And ((SchGrp.LoanBank) Is Null)) Or (((Bills.BillID)=Forms!Billing!BillingS"
    "ubForm.Form!BillID) And ((SchGrp.BLCD) Is Not Null) And ((SchGrp.Assignee_Bank) "
    "Is Not Null) And ((SchGrp.Assigned_Indicator)=\"p\")) Or (((Bills.BillID)=Forms!"
    "Billing!BillingSubForm.Form!BillID) And ((SchGrp.BLCD) Is Not Null) And ((SchGrp"
    ".Assignee_Bank) Is Null Or (SchGrp.Assignee_Bank)=\"FA llc\") And ((SchGrp.LoanB"
    "ank) Is Not Null) And ((SchGrp.SGrpFABill)=Yes))"
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
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [unittype]"
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
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="UnitsbillRent"
    Expression ="IIf(([unitstatus]=\"O\" Or [unitstatus]=\"O\" Or [unitstatus]=\"S\" Or [unitstat"
        "us]=\"r\") And [unitoffleasedt]<[billfrom],0,[unitrent])"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Alias ="TaxAmt"
    Expression ="IIf([UnitTaxExempt]=True,0,IIf(([unitstatus]=\"O\" Or [unitstatus]=\"r\" Or [uni"
        "tstatus]=\"s\") And [unitoffleasedt]<[billfrom],0,[locations].[LTaxRate]*[unitre"
        "nt]))"
    Expression ="Locations.LTaxRate"
    Expression ="Units.unittax"
    Expression ="Clients.FAVenorNo"
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
    Expression ="Bills.DueonRec"
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
        dbInteger "ColumnWidth" ="3735"
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
        dbText "Name" ="SchGrp.SGrpNegBLCD"
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
        dbText "Name" ="Clients.FAVenorNo"
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
        dbText "Name" ="SchGrp.SGrpPTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPType"
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
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssInd"
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
        dbText "Name" ="Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsbillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittax"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1243
    Bottom =773
    Left =-1
    Top =-1
    Right =1211
    Bottom =238
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-13
        Top =8
        Right =167
        Bottom =152
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =184
        Top =7
        Right =328
        Bottom =151
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =369
        Top =6
        Right =513
        Bottom =160
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =539
        Top =198
        Right =720
        Bottom =354
        Top =0
        Name ="PDstdtq"
        Name =""
    End
    Begin
        Left =305
        Top =206
        Right =492
        Bottom =369
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =124
        Top =216
        Right =268
        Bottom =360
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =533
        Top =12
        Right =677
        Bottom =143
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =892
        Top =6
        Right =1036
        Bottom =150
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =700
        Top =11
        Right =866
        Bottom =148
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1058
        Top =8
        Right =1202
        Bottom =152
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =938
        Top =213
        Right =1082
        Bottom =357
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
End
