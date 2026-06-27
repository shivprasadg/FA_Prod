Operation =1
Option =0
Where ="(((vw_SixKeys.clientID)=4971) AND ((Units.UnitStatus)<>\"S\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Type"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitGarageLocationRef"
    Expression ="Locations.LocationName"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =689
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =161
        Top =65
        Right =440
        Bottom =432
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =487
        Top =234
        Right =752
        Bottom =549
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =826
        Top =384
        Right =970
        Bottom =528
        Top =0
        Name ="Locations"
        Name =""
    End
End
