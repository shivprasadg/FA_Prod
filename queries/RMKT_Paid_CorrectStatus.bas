Operation =1
Option =0
Where ="(((Rmkt.InvoicePaidDate) Is Not Null) AND ((Units.unitstatus)=\"S\" Or (Units.un"
    "itstatus)=\"D\" Or (Units.unitstatus)=\"KS\" Or (Units.unitstatus)=\"PS\"))"
Begin InputTables
    Name ="Rmkt"
    Name ="Buyers"
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Units.unitstatus"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.FAID"
    Alias ="Buyer"
    Expression ="Buyers.ClientShNm"
    Expression ="Rmkt.SoldPrice"
End
Begin Joins
    LeftTable ="Rmkt"
    RightTable ="Buyers"
    Expression ="Rmkt.BuyerRef = Buyers.clientID"
    Flag =1
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.AssetID"
    Flag =2
    LeftTable ="Rmkt"
    RightTable ="Units"
    Expression ="Rmkt.UnitRef = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="Rmkt.InvoicePaidDate"
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
        dbText "Name" ="Units.unitstatus"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbInteger "ColumnWidth" ="2325"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbInteger "ColumnWidth" ="4380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1595
    Bottom =872
    Left =-1
    Top =-1
    Right =1571
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-3
        Top =20
        Right =191
        Bottom =389
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =220
        Top =249
        Right =399
        Bottom =397
        Top =0
        Name ="Buyers"
        Name =""
    End
    Begin
        Left =574
        Top =36
        Right =762
        Bottom =394
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =239
        Top =12
        Right =383
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
