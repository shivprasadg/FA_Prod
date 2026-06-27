Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="Rmkt"
    Name ="Buyers"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.Status"
    Alias ="Buyer"
    Expression ="Nz([clientShNm],\"none\")"
    Expression ="Rmkt.BuyOutDate"
    Expression ="Rmkt.BuyOutCost"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="vw_SixKeys.MYear"
    Alias ="Desc"
    Expression ="[Make] & \" - \" & [Model] & \" - \" & [type]"
    Alias ="BuyerId"
    Expression ="Buyers.clientID"
    Expression ="Rmkt.TargetSale"
    Expression ="Rmkt.[SUF Sent]"
    Expression ="Rmkt.[SUF Returned]"
    Expression ="Rmkt.ParkDT"
    Expression ="Units.unitoffleasedt"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Buyers"
    Expression ="Rmkt.BuyerRef = Buyers.clientID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.FAID"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
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
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyerId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.ParkDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
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
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-32
    Top =1
    Right =1265
    Bottom =804
    Left =-1
    Top =-1
    Right =1273
    Bottom =492
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =273
        Top =133
        Right =474
        Bottom =411
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =511
        Top =128
        Right =655
        Bottom =475
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =685
        Top =174
        Right =829
        Bottom =318
        Top =0
        Name ="Buyers"
        Name =""
    End
End
