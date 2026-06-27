Operation =1
Option =0
Where ="(((LeaseReport_Syndicated_V_OnOrder.OrderedYear)>=2014) And ((LeaseReport_Syndic"
    "ated_V_OnOrder.SYDYear)=0)) Or (((LeaseReport_Syndicated_V_OnOrder.OrderedYear)="
    "0) And ((LeaseReport_Syndicated_V_OnOrder.SYDYear)>=2014))"
Begin InputTables
    Name ="LeaseReport_Syndicated_V_OnOrder"
End
Begin OutputColumns
    Expression ="LeaseReport_Syndicated_V_OnOrder.*"
End
Begin OrderBy
    Expression ="LeaseReport_Syndicated_V_OnOrder.OrderedYear"
    Flag =1
    Expression ="LeaseReport_Syndicated_V_OnOrder.SYDYear"
    Flag =1
    Expression ="LeaseReport_Syndicated_V_OnOrder.UnitType"
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
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OrderedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SYDYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.LeaseReport2.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OnOrderCost"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OnOrderOEC"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SydOEC"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SydCost"
        dbLong "AggregateType" ="0"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1268
    Bottom =1021
    Left =-1
    Top =-1
    Right =1193
    Bottom =596
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =170
        Top =257
        Right =314
        Bottom =401
        Top =0
        Name ="LeaseReport_Syndicated_V_OnOrder"
        Name =""
    End
End
