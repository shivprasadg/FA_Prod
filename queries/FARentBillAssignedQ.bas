Operation =1
Option =0
Where ="(((Bills.BillID)=Forms!Billing!BillingSubForm.Form!BillID) And ((SchGrp.BLCD) Is"
    " Not Null) And ((Bills.Billto)<=Nz([unitoffleasedt],#1/1/2400#)))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Bills"
    Name ="SchGrp"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
End
Begin OutputColumns
    Expression ="Bills.BillID"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.LoanBank"
    Expression ="SchGrp.SGrpFABill"
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
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.UnitAcceptDate"
    Expression ="SchGrp.BLCD_Projected"
    Alias ="BLocST3"
    Expression ="Locations_1.LAddrLine2"
    Alias ="BLocStreet"
    Expression ="Locations_1.LocationStreet"
    Alias ="BLocST2"
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
    Alias ="FSLFee"
    Expression ="Nz([SGrpFSLPmt],0)"
    Alias ="UnitsBillRent"
    Expression ="IIf([unitoffleasedt]<[billfrom],0,IIf([Forms]![Billing]![BillingSubForm].[Form]!"
        "[UseSchRent]=-1,[UnitSRent],IIf([unitrent]=0,[UnitSRent],[UnitRent])))"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Alias ="TaxAmt"
    Expression ="IIf([UnitTaxExempt]=True,0,[Locations].[LTaxRate]*[UnitsBillRent])"
    Expression ="Locations.LTaxRate"
    Expression ="Units.unittax"
    Expression ="Clients.FAVenorNo"
    Expression ="Units.unitoffleasedt"
    Alias ="Bill1"
    Expression ="[invfld1desc] & \" \" & [invfld1]"
    Alias ="Bill2"
    Expression ="[invfld2desc] & \" \" & [invfld2] & \" \""
    Expression ="Bills.DueonRec"
    Expression ="SchGrp.AssignmentDate"
    Expression ="Schedule.invfld1desc"
    Expression ="Schedule.invfld1"
    Expression ="Schedule.invfld2desc"
    Expression ="Schedule.invfld2"
    Expression ="Clients.clientID"
    Alias ="RentNoTax"
    Expression ="[UnitsbillRent]"
    Alias ="FSLTax"
    Expression ="IIf([UnitTaxExempt]=False,Nz([FSLFee],0)*Nz([locations].[LTaxRate],0),0)"
    Alias ="UnitTotalRent"
    Expression ="[UnitsBillRent]+Nz([FSLFee],0)+Nz([TaxAmt],0)"
    Alias ="BillToAddress"
    Expression ="[BLocStreet] & Chr(13) & Chr(10) & [BLocSt2] & Chr(13) & Chr(10) & IIf(IsNull([B"
        "LocST3]),Null,[BLocST3] & Chr(13) & Chr(10)) & [blOCcITY] & \", \" & [bLOCST] & "
        "\" \" & Left([BlOCzIP],5) & \"-\" & Right([BlOCzIP],4)"
    Alias ="InvoiceNumber"
    Expression ="Bills.BillNo"
    Alias ="DueDate"
    Expression ="IIf([DueonRec]=-1,\"On Receipt\",IIf(IsNull([BillDueDt]),[billFrom],[BillDueDt])"
        ")"
    Alias ="BankContract"
    Expression ="[SchBnkClNO] & IIf(Nz([SgrpBnkNo],\"\")=\"\",Null,\"-\" & [SgrpBnkNo])"
    Alias ="GroupID"
    Expression ="SchGrp.SGrpID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
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
    Begin
        dbText "Name" ="Schedule.invfld1desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.invfld1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.invfld2desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.invfld2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentNoTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DueDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
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
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankContract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
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
        dbText "Name" ="Expr1049"
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
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =15
    Top =272
    Right =1222
    Bottom =1147
    Left =-1
    Top =-1
    Right =1183
    Bottom =273
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
        Left =497
        Top =203
        Right =674
        Bottom =366
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =192
        Top =212
        Right =336
        Bottom =356
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
        Bottom =353
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
End
