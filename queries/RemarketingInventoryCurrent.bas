Operation =1
Option =0
Where ="(((RemarketingReportAllUnits.UnitStatus)=\"R\" Or (RemarketingReportAllUnits.Uni"
    "tStatus)=\"PS\"))"
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.[Sch|Grp]"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.UnitDescShortVer"
    Expression ="RemarketingReportAllUnits.ODORead"
    Expression ="RemarketingReportAllUnits.ODOReadDate"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.DaysInInventory"
    Expression ="RemarketingReportAllUnits.ParkDateProjected"
    Expression ="RemarketingReportAllUnits.OffLeaseDate"
    Expression ="RemarketingReportAllUnits.InvoiceSent"
    Expression ="RemarketingReportAllUnits.Parked_v_EOL_Days"
    Expression ="RemarketingReportAllUnits.Residual"
    Expression ="RemarketingReportAllUnits.TermPenalty"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.BuyOut"
    Expression ="RemarketingReportAllUnits.FMVBase"
    Expression ="RemarketingReportAllUnits.FMV"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Expression ="RemarketingReportAllUnits.IFDate"
    Expression ="RemarketingReportAllUnits.IFAmount"
    Expression ="RemarketingReportAllUnits.IFBank"
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
        dbText "Name" ="RemarketingReportAllUnits.SourceCompany"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Parked_v_EOL_Days"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2265"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ODORead"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ODOReadDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[Sch|Grp]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3450"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Residual"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.OffLeaseDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.IFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDateProjected"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TermPenalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.IFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMVBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.IFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.DaysInInventory"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceSent"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1684
    Bottom =994
    Left =-1
    Top =-1
    Right =1395
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =358
        Top =27
        Right =502
        Bottom =566
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
