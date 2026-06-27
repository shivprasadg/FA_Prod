Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"TR\" Or (vw_SixKeys.Status)=\"R\" Or (vw_SixKeys.Status)"
    "=\"S\" Or (vw_SixKeys.Status)=\"K\" Or (vw_SixKeys.Status)=\"KS\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Rmkt"
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.Status"
    Alias ="SourceCompany"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Assignee"
    Expression ="Nz([BankShortName],\"N/A\")"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Type"
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
    Expression ="Rmkt.MiParked"
    Expression ="Rmkt.InspMileage1"
    Expression ="Locations.LocationName"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Rmkt"
    Expression ="vw_SixKeys.AssetID = Rmkt.UnitRef"
    Flag =1
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.MiParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspMileage1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Model"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =37
    Top =-23
    Right =1526
    Bottom =975
    Left =-1
    Top =-1
    Right =1457
    Bottom =698
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =90
        Top =105
        Right =401
        Bottom =678
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =973
        Top =248
        Right =1117
        Bottom =617
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =438
        Top =140
        Right =608
        Bottom =631
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =670
        Top =96
        Right =814
        Bottom =664
        Top =0
        Name ="Units"
        Name =""
    End
End
