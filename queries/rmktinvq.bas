Operation =1
Option =0
Where ="(((Rmkt.InvoiceRequestDate)=forms!ReMarketInvoiceManager!reqdt) And ((Locations."
    "LocationType)=\"billing\") And ((Rmkt.BuyerRef)=forms!ReMarketInvoiceManager!buy"
    "erref))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Rmkt.InvoiceRequested"
    Expression ="Rmkt.InvoiceRequestDate"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.InvoiceSentDate"
    Expression ="Clients.clientCompanyName"
    Expression ="Locations.LocationType"
    Expression ="Locations.*"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
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
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PInvReq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AInvREq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.rmktinvsent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =116
        Top =-4
        Right =333
        Bottom =140
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =371
        Top =0
        Right =597
        Bottom =186
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =658
        Top =11
        Right =802
        Bottom =155
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =850
        Top =12
        Right =994
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End
