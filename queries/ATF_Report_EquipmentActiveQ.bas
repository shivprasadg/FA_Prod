Operation =1
Option =1
Where ="SchID=2958"
Begin InputTables
    Name ="ATF_Report_EquipmentActive"
End
Begin OutputColumns
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
        dbText "Name" ="ATF_Report_EquipmentActive.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.GroupOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.SpecProfile"
        dbInteger "ColumnWidth" ="6450"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.ResidualSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.UnitRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.PartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.ResidualSynd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.UnitCnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_EquipmentActive.GroupsIn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-357
    Top =104
    Right =1240
    Bottom =854
    Left =-1
    Top =-1
    Right =1579
    Bottom =136
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ATF_Report_EquipmentActive"
        Name =""
    End
End
