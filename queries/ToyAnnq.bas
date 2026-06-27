Operation =1
Option =0
Where ="(((Schedule.SchID)=Forms!Billing!SchID) And ((SchGrp.BLCD) Is Not Null) And ((Un"
    "its.UnitInServ) Is Not Null) And ((Locations_1.LocationType)=\"billing\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="lookup_States"
    Name ="Locations"
    Alias ="Locations_1"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.BLCD"
    Alias ="DailyServ"
    Expression ="IIf([UnitType]=\"Tractor\",Round(0.0329*[sgrpfslpmt],2),0)"
    Alias ="DailyNon-Serv"
    Expression ="IIf([UnitType]=\"Tractor\",Round(0.0329*[sgrprsltpmt],2),0)"
    Expression ="Units.UnitInServ"
    Expression ="Units.UnitDocAccpt"
    Expression ="Units.UnitInServInd"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Alias ="BLocStreet"
    Expression ="Locations_1.LocationStreet"
    Alias ="BLOCst2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="BLOCST3"
    Expression ="Locations_1.LAddrLine2"
    Alias ="BLOCCity"
    Expression ="Locations_1.LCity"
    Alias ="BLOCST"
    Expression ="Locations_1.LState"
    Alias ="BLOCZip"
    Expression ="Locations.LZip"
    Expression ="Clients.FAVenorNo"
    Alias ="Bill1"
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    Alias ="Bill2"
    Expression ="UCase([invfld2desc] & \" \" & IIf([invfld2]=\"0\",Null,[invfld2]))"
    Expression ="Units.unitunitnum"
    Expression ="SchGrp.SGRPNoPD"
    Expression ="Units.UnitTaxExempt"
    Expression ="Units.unitmake"
    Expression ="MstrLease.MLNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationID"
    Expression ="Locations.LocationName"
    Expression ="Locations.LTaxRate"
    Expression ="lookup_States.StTaxType"
    Alias ="InvoiceNumber"
    Expression ="[Forms]![Billing]![BillingSubForm].[Form]![BillNo]"
    Alias ="invfrdt"
    Expression ="[Forms]![Billing]![BillingSubForm].[Form]![billfrom]"
    Alias ="invtodt"
    Expression ="[Forms]![Billing]![BillingSubForm].[Form]![billto]"
    Alias ="invduedt"
    Expression ="[Forms]![Billing]![BillingSubForm].[Form]![billduedt]"
    Alias ="invdt"
    Expression ="[Forms]![Billing]![BillingSubForm].[Form]![billdt]"
    Expression ="Schedule.SchFSL"
    Expression ="SchGrp.sgrprsltpmt"
    Expression ="Units.unitvin"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitType"
    Expression ="Units.unitpdrent"
    Expression ="SchGrp.SGrpID"
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
    LeftTable ="Clients"
    RightTable ="Locations_1"
    Expression ="Clients.ClientGroupId = Locations_1.ClientGroupID"
    Flag =1
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DailyServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DailyNon-Serv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BLocST"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLOCst2"
        dbInteger "ColumnWidth" ="4455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLOCST3"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLOCCity"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocStreet"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
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
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoCZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="invfrdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="invtodt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="invduedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="invdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrprsltpmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitpdrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1025"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1234
    Bottom =308
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =323
        Top =142
        Right =508
        Bottom =294
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =196
        Top =12
        Right =340
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =355
        Top =7
        Right =499
        Bottom =151
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =512
        Top =11
        Right =655
        Bottom =155
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =685
        Top =9
        Right =867
        Bottom =153
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =928
        Top =10
        Right =1084
        Bottom =203
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =1117
        Top =24
        Right =1261
        Bottom =168
        Top =0
        Name ="lookup_States"
        Name =""
    End
End
