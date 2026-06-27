Operation =1
Option =0
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Alias ="Account Name"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.Schedule"
    Expression ="RemarketingReportAllUnits.UnitGroup"
    Alias ="Cust Unit#"
    Expression ="RemarketingReportAllUnits.Unitnum"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.Odometer"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.ParkDateProjected"
    Expression ="RemarketingReportAllUnits.InspectedDate"
    Expression ="RemarketingReportAllUnits.LocationName"
    Expression ="RemarketingReportAllUnits.LocationAddress"
    Expression ="RemarketingReportAllUnits.LCity"
    Expression ="RemarketingReportAllUnits.LState"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Alias ="Actual Sale Price"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Expression ="RemarketingReportAllUnits.GLTarget"
    Alias ="VarPct"
    Expression ="RemarketingReportAllUnits.GLTargetMargin"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Alias ="Days Since BuyOut"
    Expression ="RemarketingReportAllUnits.DaysInInventoryPerAccounting"
    Expression ="RemarketingReportAllUnits.ClientID"
    Expression ="RemarketingReportAllUnits.BuyerID"
    Expression ="RemarketingReportAllUnits.UnitID"
End
Begin OrderBy
    Expression ="RemarketingReportAllUnits.SourceCompany"
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
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cust Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Actual Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InspectedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.GLTarget"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VarPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDateProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOutDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.DaysInInventoryPerAccounting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =7
    Top =79
    Right =1147
    Bottom =1060
    Left =-1
    Top =-1
    Right =1108
    Bottom =502
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =337
        Top =22
        Right =572
        Bottom =480
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
