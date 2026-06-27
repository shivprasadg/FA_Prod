Operation =1
Option =8
Where ="(((AllLeasedUnits_Cost.SchID)=[Forms]![ScheduleControl]![Schid]))"
Begin InputTables
    Name ="AllLeasedUnits_Cost"
End
Begin OutputColumns
    Expression ="AllLeasedUnits_Cost.Schedule"
    Expression ="AllLeasedUnits_Cost.Unitnum"
    Expression ="AllLeasedUnits_Cost.VIN"
    Alias ="SCH_Grp"
    Expression ="AllLeasedUnits_Cost.UnitGroup"
    Expression ="AllLeasedUnits_Cost.UnitCost"
    Expression ="AllLeasedUnits_Cost.OBCMake"
    Expression ="AllLeasedUnits_Cost.OBCInvoice"
    Expression ="AllLeasedUnits_Cost.OBCCost"
    Expression ="AllLeasedUnits_Cost.OBCInstaller"
    Expression ="AllLeasedUnits_Cost.OBCInstall"
    Expression ="AllLeasedUnits_Cost.OBCInstall_Invoice"
    Expression ="AllLeasedUnits_Cost.WarrantyCost"
    Expression ="AllLeasedUnits_Cost.OtherCost"
    Expression ="AllLeasedUnits_Cost.FeesTaxesCost"
    Alias ="TotalCost"
    Expression ="[UnitCost]+[obcinstall]+[WarrantyCost]+[OtherCost]+[FeesTaxesCost]"
End
Begin OrderBy
    Expression ="AllLeasedUnits_Cost.VIN"
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
        dbText "Name" ="AllLeasedUnits_Cost.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="6225"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCInstaller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCInstall_Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OBCCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SCH_Grp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.WarrantyCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.OtherCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Totalcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AllLeasedUnits_Cost.FeesTaxesCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =105
    Top =37
    Right =1692
    Bottom =869
    Left =-1
    Top =-1
    Right =1555
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =107
        Top =93
        Right =324
        Bottom =520
        Top =0
        Name ="AllLeasedUnits_Cost"
        Name =""
    End
End
