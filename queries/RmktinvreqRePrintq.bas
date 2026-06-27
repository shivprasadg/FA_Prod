Operation =1
Option =0
Where ="(((Rmkt.BuyerRef)=forms!ReMarketInvoiceManager!buyerref) And ((Rmkt.InvoiceReque"
    "stDate)=forms!ReMarketInvoiceManager!RequestDate) And ((RmktBills.ID)=forms!ReMa"
    "rketInvoiceManager!id))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
    Name ="Locations"
    Name ="lookup_States"
    Name ="RmktBills"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Rmkt.RmktId"
    Expression ="Rmkt.BuyerRef"
    Expression ="Clients.clientID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientCoNm2"
    Expression ="Clients.clientDBA"
    Expression ="Clients.clientAddress1"
    Expression ="Clients.clientAddress2"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientZip"
    Expression ="Clients.clientZip"
    Expression ="Clients.[Dlr/Resale]"
    Expression ="Units.unitvin"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Rmkt.SoldPrice"
    Expression ="Clients.clientState"
    Expression ="Rmkt.TargetSale"
    Expression ="Rmkt.InvoiceRequested"
    Expression ="Units.unitstatus"
    Expression ="Rmkt.RmkSm"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.InvoiceRequestDate"
    Expression ="Locations.*"
    Expression ="Units.unitvin"
    Alias ="srt"
    Expression ="[LCity] & \" \" & [LCounty] & \" \" & [LState]"
    Expression ="lookup_States.StTaxType"
    Expression ="Clients.clientZip"
    Expression ="Clients.clientZip"
    Expression ="RmktBills.Billdt"
    Expression ="RmktBills.RBillDueDt"
    Expression ="Rmkt.InvoiceRequestDate"
    Expression ="RmktBills.RevNo"
    Expression ="RmktBills.ID"
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
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="lookup_States"
    Expression ="Clients.clientState = lookup_States.StateAbbrev"
    Flag =1
    LeftTable ="Clients"
    RightTable ="RmktBills"
    Expression ="Clients.clientID = RmktBills.BuyerRef"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[Dlr/Resale]"
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
        dbText "Name" ="Rmkt.rmktsalepr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.rmktsaledt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PInvReq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AInvREq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.District"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSubLessee"
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
        dbText "Name" ="Locations.LocationCounty"
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
        dbText "Name" ="Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locationPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locationFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocDCs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocDCMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSlprs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locSplrMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrucks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrkMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrl"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrlMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocReeferTrlrs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocRefrigUnitHPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsCmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.oldlocid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.RLocId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTaxOpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCostCtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocBKACCT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LOCCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="srt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1011"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1031"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.RBillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.RevNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1026"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ainvreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-249
    Top =52
    Right =1009
    Bottom =698
    Left =-1
    Top =-1
    Right =1240
    Bottom =252
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =56
        Top =3
        Right =200
        Bottom =147
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =285
        Top =13
        Right =429
        Bottom =157
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =515
        Top =10
        Right =659
        Bottom =153
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =26
        Top =159
        Right =170
        Bottom =303
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =746
        Top =12
        Right =890
        Bottom =156
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =695
        Top =137
        Right =881
        Bottom =372
        Top =0
        Name ="RmktBills"
        Name =""
    End
End
