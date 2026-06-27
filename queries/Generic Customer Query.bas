Operation =1
Option =0
Having ="(((Clients.ClientShNm) Like \"*Transc*\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.unitvin"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.FundDate_Projected"
    Expression ="SchGrp.UnitsInGroup"
    Expression ="SchGrp.OrigEquipCost"
    Expression ="SchGrp.LeaseTermPrimary"
    Alias ="Expr1"
    Expression ="SchGrp.SGrpPLRF"
    Expression ="SchGrp.LeasePymtForGroup"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="SchGrp.SGrpFSLSch"
    Expression ="SchGrp.ExchangeITGroup"
    Expression ="SchGrp.ExchangeITMonth"
    Expression ="SchGrp.ExchangeITType"
    Expression ="SchGrp.LeaseTermNotice"
    Expression ="SchGrp.LeaseExtOption1Term"
    Expression ="SchGrp.LeaseExtOption1LRF"
    Expression ="SchGrp.SGrpRtTBD"
    Expression ="SchGrp.SGrpOpt2"
    Expression ="SchGrp.SGrpO2LRF"
    Expression ="SchGrp.SGrpOpt3"
    Expression ="SchGrp.SGrpO3LRF"
    Expression ="SchGrp.SGrpEngReb"
    Expression ="SchGrp.SgrpEngMi"
    Expression ="SchGrp.SGrpRefR"
    Expression ="SchGrp.SGrpRHr"
    Expression ="SchGrp.AllowedMilesPer"
    Expression ="SchGrp.SGprAllow1"
    Expression ="SchGrp.SGrpMH1"
    Expression ="SchGrp.SGrpov1"
    Expression ="SchGrp.SGprAllow2"
    Expression ="SchGrp.SGrpMH2"
    Expression ="SchGrp.SGrpov2"
    Expression ="SchGrp.SGprAllow3"
    Expression ="SchGrp.SGrpMH3"
    Expression ="SchGrp.SGrpov3"
    Expression ="SchGrp.SGprAllow4"
    Expression ="SchGrp.SGrpov4"
    Expression ="SchGrp.SGrpMH4"
    Expression ="SchGrp.AllowedMilesPeriod"
    Alias ="Desc"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unittype] & \" \" & [unitmodel]"
    Expression ="Units.unitstatus"
    Expression ="SchGrp.RebateUnitProjected"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin OrderBy
    Expression ="Units.unitvin"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
    Expression ="SchGrp.UnitsInGroup"
    Flag =0
End
Begin Groups
    Expression ="Units.unitvin"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.FundDate"
    GroupLevel =0
    Expression ="SchGrp.FundDate_Projected"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="SchGrp.OrigEquipCost"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.SGrpPLRF"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtForGroup"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLSch"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITGroup"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITMonth"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITType"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermNotice"
    GroupLevel =0
    Expression ="SchGrp.LeaseExtOption1Term"
    GroupLevel =0
    Expression ="SchGrp.LeaseExtOption1LRF"
    GroupLevel =0
    Expression ="SchGrp.SGrpRtTBD"
    GroupLevel =0
    Expression ="SchGrp.SGrpOpt2"
    GroupLevel =0
    Expression ="SchGrp.SGrpO2LRF"
    GroupLevel =0
    Expression ="SchGrp.SGrpOpt3"
    GroupLevel =0
    Expression ="SchGrp.SGrpO3LRF"
    GroupLevel =0
    Expression ="SchGrp.SGrpEngReb"
    GroupLevel =0
    Expression ="SchGrp.SgrpEngMi"
    GroupLevel =0
    Expression ="SchGrp.SGrpRefR"
    GroupLevel =0
    Expression ="SchGrp.SGrpRHr"
    GroupLevel =0
    Expression ="SchGrp.AllowedMilesPer"
    GroupLevel =0
    Expression ="SchGrp.SGprAllow1"
    GroupLevel =0
    Expression ="SchGrp.SGrpMH1"
    GroupLevel =0
    Expression ="SchGrp.SGrpov1"
    GroupLevel =0
    Expression ="SchGrp.SGprAllow2"
    GroupLevel =0
    Expression ="SchGrp.SGrpMH2"
    GroupLevel =0
    Expression ="SchGrp.SGrpov2"
    GroupLevel =0
    Expression ="SchGrp.SGprAllow3"
    GroupLevel =0
    Expression ="SchGrp.SGrpMH3"
    GroupLevel =0
    Expression ="SchGrp.SGrpov3"
    GroupLevel =0
    Expression ="SchGrp.SGprAllow4"
    GroupLevel =0
    Expression ="SchGrp.SGrpov4"
    GroupLevel =0
    Expression ="SchGrp.SGrpMH4"
    GroupLevel =0
    Expression ="SchGrp.AllowedMilesPeriod"
    GroupLevel =0
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unittype] & \" \" & [unitmodel]"
    GroupLevel =0
    Expression ="Units.unitstatus"
    GroupLevel =0
    Expression ="SchGrp.RebateUnitProjected"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitUnitNum"
    GroupLevel =0
    Expression ="Schedule.SchID"
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1005"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnWidth" ="1155"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoUnits"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpEquipCost"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPTerm"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPLRF"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPRent"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpOpt1"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO1LRF"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpOpt2"
        dbInteger "ColumnOrder" ="33"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpEngReb"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEngMi"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow1"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow2"
        dbInteger "ColumnOrder" ="34"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH1"
        dbInteger "ColumnOrder" ="31"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH2"
        dbInteger "ColumnOrder" ="35"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH3"
        dbInteger "ColumnOrder" ="36"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH4"
        dbInteger "ColumnOrder" ="37"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov1"
        dbInteger "ColumnOrder" ="32"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.[SGPRU/A]"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpUT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLSch"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExch"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExMo"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRtTBD"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="42"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="43"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpExType"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpExtNot"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRHr"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssignee"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPrjFund"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFunddt"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRefR"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO2LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpOpt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO3LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.[SGRPREBATE/UNIT]"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
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
    Right =1226
    Bottom =244
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
        Left =178
        Top =14
        Right =322
        Bottom =158
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =300
        Top =14
        Right =444
        Bottom =158
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =422
        Top =13
        Right =566
        Bottom =157
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =582
        Top =12
        Right =739
        Bottom =262
        Top =0
        Name ="Units"
        Name =""
    End
End
