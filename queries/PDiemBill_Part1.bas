Operation =1
Option =0
Having ="(((Bills.BillID)=Forms!Billing!BillingSubForm.Form!BillID) And ((SchGrp.SGrpNoSl"
    "v) Is Not Null) And ((SchGrp.SGRPNoPD) Is Not Null))"
Begin InputTables
    Name ="Bills"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Bills.BillID"
    Expression ="Clients.clientID"
    Expression ="MstrLease.MLID"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
    Expression ="Units.UnitID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.SGrpNoSlv"
    Expression ="SchGrp.SGRPNoPD"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitAcceptDate"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [Un"
        "itSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [UnitType])"
    Expression ="Units.UnitTaxExempt"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.FAVenorNo"
    Expression ="Schedule.SchBnkClNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.unitpdrent"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.unitinserv"
    Expression ="Schedule.schcrexp"
    Expression ="Schedule.SchFSL"
    Expression ="Units.unitrent"
    Expression ="Units.unittax"
    Expression ="Units.UNITOFFLEASEDT"
    Expression ="Units.UnitPDRent"
    Expression ="Schedule.SchCurrRent"
    Expression ="Bills.BillDueDt"
    Expression ="Bills.DueonRec"
    Expression ="Bills.Billdt"
    Expression ="Bills.BillSent"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Bills"
    RightTable ="Schedule"
    Expression ="Bills.BillSchRef = Schedule.SchID"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
End
Begin Groups
    Expression ="Bills.BillID"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.SGrpNoSlv"
    GroupLevel =0
    Expression ="SchGrp.SGRPNoPD"
    GroupLevel =0
    Expression ="Units.unitunitnum"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [Un"
        "itSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [UnitType])"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.FAVenorNo"
    GroupLevel =0
    Expression ="Schedule.SchBnkClNO"
    GroupLevel =0
    Expression ="Right([unitvin],6)"
    GroupLevel =0
    Expression ="Units.unitpdrent"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="Units.unitinserv"
    GroupLevel =0
    Expression ="Schedule.schcrexp"
    GroupLevel =0
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="Units.unitrent"
    GroupLevel =0
    Expression ="Units.unittax"
    GroupLevel =0
    Expression ="Units.UNITOFFLEASEDT"
    GroupLevel =0
    Expression ="Units.UnitPDRent"
    GroupLevel =0
    Expression ="Schedule.SchCurrRent"
    GroupLevel =0
    Expression ="Bills.BillDueDt"
    GroupLevel =0
    Expression ="Bills.DueonRec"
    GroupLevel =0
    Expression ="Bills.Billdt"
    GroupLevel =0
    Expression ="Bills.BillSent"
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
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="2835"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2130"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnWidth" ="1875"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbInteger "ColumnWidth" ="1320"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2985"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitpdrent"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoSlv"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.schcrexp"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITOFFLEASEDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.unitcurrentpd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.[SchCurrRent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchCurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitPDRent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.[SgrpNoPD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[unittax]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillSent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =163
    Top =79
    Right =1421
    Bottom =909
    Left =-1
    Top =-1
    Right =1226
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =7
        Top =6
        Right =151
        Bottom =150
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =350
        Top =147
        Right =518
        Bottom =377
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =6
        Top =160
        Right =150
        Bottom =388
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =181
        Top =89
        Right =325
        Bottom =243
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =560
        Top =71
        Right =701
        Bottom =215
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =731
        Top =74
        Right =897
        Bottom =244
        Top =0
        Name ="Units"
        Name =""
    End
End
