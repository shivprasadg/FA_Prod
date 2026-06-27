Operation =1
Option =0
Where ="(((Bills.BillID)=Forms!Billing!BillID) And ((SchGrp.BLCD) Is Not Null) And ((Uni"
    "ts.unitoffleasedt) Is Null Or (Units.unitoffleasedt)>=bills.Billfrom))"
Begin InputTables
    Name ="Bills"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="lookup_States"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Banks"
End
Begin OutputColumns
    Expression ="Bills.BillID"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.BLCD"
    Alias ="BillingAddress"
    Expression ="[locations_1].[LocationStreet] & Chr(13) & Chr(10) & IIf(IsNull([locations_1].[L"
        "AddrLine1]),Null,[locations_1].[LAddrLine1] & Chr(13) & Chr(10)) & IIf(IsNull([l"
        "ocations_1].[LAddrLine2]),Null,[locations_1].[LAddrLine2] & Chr(13) & Chr(10)) &"
        " [locations_1].[LCity] & \" \" & [locations_1].[LState] & \" \" & [locations_1]."
        "[LZip]"
    Alias ="Serv"
    Expression ="IIf([UnitType]=\"Tractor\",[Sgrpfslpmt],0)"
    Alias ="Non-Serv"
    Expression ="IIf([UnitType]=\"Tractor\",[sgrprsltpmt],0)"
    Expression ="Units.UnitInServ"
    Expression ="Units.UnitInServInd"
    Expression ="Clients.FAVenorNo"
    Alias ="Bill1"
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    Alias ="Bill2"
    Expression ="UCase([invfld2desc] & \" \" & IIf([invfld2]=\"0\",Null,[invfld2]))"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitTaxExempt"
    Expression ="Units.unitmake"
    Expression ="MstrLease.MLNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationID"
    Expression ="Locations.LocationName"
    Expression ="Locations.LState"
    Expression ="Locations.LCounty"
    Expression ="Locations.LCity"
    Expression ="Locations.LTaxRate"
    Expression ="lookup_States.StTaxType"
    Alias ="InvoiceNumber"
    Expression ="Bills.billno"
    Alias ="invfrdt"
    Expression ="Bills.Billfrom"
    Alias ="invtodt"
    Expression ="Bills.Billto"
    Alias ="invduedt"
    Expression ="Bills.BillDueDt"
    Alias ="invdt"
    Expression ="Bills.Billdt"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="SchGrp.sgrprsltpmt"
    Expression ="Units.UNITRENT"
    Expression ="Units.unitmodel"
    Expression ="Units.UnitVIN"
    Expression ="SchGrp.AssignmentDate"
    Expression ="Banks.BankName"
    Expression ="Units.unitoffleasedt"
    Alias ="UnitLineTotal"
    Expression ="IIf([UnitRent]<1,0,[serv]+[non-serv]+Round(IIf([UnitTaxExempt]=Yes,0,[non-serv]*"
        "[Locations].[LTaxRate]),2))"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Bills"
    RightTable ="Schedule"
    Expression ="Bills.BillSchRef = Schedule.SchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Locations_1"
    Expression ="SchGrp.BillToAddressID = Locations_1.LocationID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Banks"
    Expression ="SchGrp.Assignee_Bank = Banks.BankShortName"
    Flag =1
End
Begin OrderBy
    Expression ="Units.unitunitnum"
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
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Bill1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Bill2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="invfrdt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="invtodt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="invduedt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="invdt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="SchGrp.sgrprsltpmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Serv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Non-Serv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Clients_1.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLineTotal"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankName"
        dbInteger "ColumnWidth" ="3885"
        dbInteger "ColumnOrder" ="37"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-32
    Top =34
    Right =1585
    Bottom =985
    Left =-1
    Top =-1
    Right =1593
    Bottom =366
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =301
        Top =153
        Right =445
        Bottom =297
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =-2
        Top =6
        Right =142
        Bottom =150
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =151
        Top =5
        Right =295
        Bottom =149
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =450
        Bottom =150
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =462
        Top =8
        Right =618
        Bottom =301
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =685
        Top =9
        Right =842
        Bottom =195
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =981
        Top =1
        Right =1137
        Bottom =194
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =1171
        Top =205
        Right =1315
        Bottom =349
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =466
        Top =303
        Right =611
        Bottom =456
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =939
        Top =99
        Right =1083
        Bottom =325
        Top =0
        Name ="Banks"
        Name =""
    End
End
