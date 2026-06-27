Operation =1
Option =0
Begin InputTables
    Name ="GroupFinalizationReport_base"
End
Begin OutputColumns
    Expression ="GroupFinalizationReport_base.Client"
    Expression ="GroupFinalizationReport_base.Schedule"
    Expression ="GroupFinalizationReport_base.UnitGroup"
    Alias ="Assigned"
    Expression ="GroupFinalizationReport_base.Assigned_Indicator"
    Expression ="GroupFinalizationReport_base.AssignmentDate"
    Alias ="SyndFeeProj"
    Expression ="IIf([GroupFinalizationReport_base].[Fee_Projected]=0,DLookUp(\"Person\",\"GroupF"
        "inalizedPersons\",\"Fieldname='Fee_Projected'\"))"
    Alias ="SynFee"
    Expression ="IIf([GroupFinalizationReport_base].[SynFee]=0,DLookUp(\"Person\",\"GroupFinalize"
        "dPersons\",\"Fieldname='SynFee'\"))"
    Alias ="Index(LO)"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SGrpSalesIndex]),DLookUp(\"Person\",\""
        "GroupFinalizedPersons\",\"Fieldname='SGrpSalesIndex'\"))"
    Alias ="Index"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SGrpIndex]),DLookUp(\"Person\",\"Grou"
        "pFinalizedPersons\",\"Fieldname='SGrpIndex'\"))"
    Alias ="Yield(LO)"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SgrpSalesYield]),DLookUp(\"Person\",\""
        "GroupFinalizedPersons\",\"Fieldname='SgrpSalesYield'\"))"
    Alias ="Yield"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SgrpYield]),DLookUp(\"Person\",\"Grou"
        "pFinalizedPersons\",\"Fieldname='SgrpYield'\"))"
    Alias ="RAL Back"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SGrpRALB]),DLookUp(\"Person\",\"Group"
        "FinalizedPersons\",\"Fieldname='SGrpRALB'\"))"
    Alias ="SchIn"
    Expression ="IIf(IsNull([GroupFinalizationReport_base].[SchIn]),DLookUp(\"Person\",\"GroupFin"
        "alizedPersons\",\"Fieldname='SchIn'\"))"
    Alias ="PDUnit"
    Expression ="GroupFinalizationReport_base.PDiemRent"
    Expression ="GroupFinalizationReport_base.BillToAddress"
    Expression ="GroupFinalizationReport_base.BLCDBasis"
    Expression ="GroupFinalizationReport_base.UnitSyndRent"
    Expression ="GroupFinalizationReport_base.UnitResidual"
    Expression ="GroupFinalizationReport_base.SchID"
    Expression ="GroupFinalizationReport_base.SGrpID"
End
Begin OrderBy
    Expression ="GroupFinalizationReport_base.Client"
    Flag =0
    Expression ="GroupFinalizationReport_base.Schedule"
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
dbMemo "Filter" ="([GroupFinalizationReport_A].[Assigned]<>\"F\" Or [GroupFinalizationReport_A].[A"
    "ssigned] IS Null)"
Begin
    Begin
        dbText "Name" ="GroupFinalizationReport_base.Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.UnitGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="SynFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="SchIn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Assigned"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Index"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="SyndFeeProj"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Index(LO)"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Yield(LO)"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Yield"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="RAL Back"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="PDUnit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.BillToAddress"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.UnitResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.BLCDBasis"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.UnitSyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.AssignmentDate"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizationReport_base.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =56
    Right =1391
    Bottom =798
    Left =-1
    Top =-1
    Right =1321
    Bottom =377
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =190
        Top =23
        Right =334
        Bottom =305
        Top =0
        Name ="GroupFinalizationReport_base"
        Name =""
    End
End
