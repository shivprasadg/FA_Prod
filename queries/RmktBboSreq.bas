Operation =1
Option =0
Where ="(((Rmkt.InvoiceRequestNum)=forms!ReMarketInvoiceManager!InvoiceRequestID) And (("
    "Rmkt.InvoicePaidDate) Is Not Null) And ((Rmkt.pbos)=True))"
Begin InputTables
    Name ="Clients"
    Alias ="Clients_1"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
    Name ="Locations"
    Name ="LookUp_UnitDesc"
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
    Expression ="Clients.clientState"
    Alias ="Buyerzip"
    Expression ="Clients.clientZip"
    Expression ="Clients.[Dlr/Resale]"
    Expression ="Units.unitvin"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.TargetSale"
    Expression ="Rmkt.InvoiceRequested"
    Expression ="Units.unitstatus"
    Expression ="Rmkt.RmkSm"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.InvoiceRequestNum"
    Expression ="Rmkt.InvoiceRequestDate"
    Expression ="Rmkt.InvoicePaidDate"
    Alias ="SellerCoNm"
    Expression ="Clients_1.clientCompanyName"
    Alias ="SellerCo2"
    Expression ="Clients_1.clientCoNm2"
    Alias ="SellerDBA"
    Expression ="Clients_1.clientDBA"
    Alias ="SellerAdd1"
    Expression ="Clients_1.clientAddress1"
    Alias ="SellerAdd2"
    Expression ="Clients_1.clientAddress2"
    Alias ="SellerCity"
    Expression ="Clients_1.clientCity"
    Alias ="SellerSt"
    Expression ="Clients_1.clientState"
    Alias ="SellerZip"
    Expression ="Clients_1.clientZip"
    Expression ="Rmkt.SoldPrice"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Rmkt.pbos"
    Alias ="Buyer"
    Expression ="\"<Strong>\" & [clients].[clientCompanyName] & \"</Strong><br>\" & IIf(IsNull([c"
        "lients].[clientCoNm2]),Null,[clients].[clientCoNm2] & \"<br>\") & IIf(IsNull([cl"
        "ients].[clientDBA]),Null,[clients].[clientDBA] & \"<br>\") & IIf(IsNull([clients"
        "].[clientAddress1]),Null,[clients].[clientAddress1] & \"<br>\") & IIf(IsNull([cl"
        "ients].[clientAddress2]),Null,[clients].[clientAddress2] & \"<br>\") & [clients]"
        ".[clientcity] & \" \" & [clients].[clientState] & \" \" & [clients].[clientZip]"
    Alias ="SalePrice"
    Expression ="CCur(Nz([rmkt].[SoldPrice],0))"
    Alias ="Client"
    Expression ="\"<Strong>\" & [clients_1].[clientCompanyName] & \"</Strong><br>\" & IIf(IsNull("
        "[clients_1].[clientCoNm2]),Null,[clients_1].[clientCoNm2] & \"<br>\") & IIf(IsNu"
        "ll([clients_1].[clientDBA]),Null,[clients_1].[clientDBA] & \"<br>\") & IIf(IsNul"
        "l([clients_1].[clientAddress1]),Null,[clients_1].[clientAddress1] & \"<br>\") & "
        "IIf(IsNull([clients_1].[clientAddress2]),Null,[clients_1].[clientAddress2] & \"<"
        "br>\") & [clients_1].[clientcity] & \" \" & [clients_1].[clientstate] & \" \" & "
        "[clients_1].[clientZip]"
    Expression ="LookUp_UnitDesc.DescriptionMed"
End
Begin Joins
    LeftTable ="Clients_1"
    RightTable ="MstrLease"
    Expression ="Clients_1.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
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
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="LookUp_UnitDesc"
    Expression ="Units.UnitID = LookUp_UnitDesc.AssetID"
    Flag =1
End
Begin OrderBy
    Expression ="Units.unitvin"
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
        dbInteger "ColumnWidth" ="2265"
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
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerCoNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerCo2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerAdd1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerAdd2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SellerZip"
        dbLong "AggregateType" ="-1"
        dbMemo "Caption" ="sellerzip"
    End
    Begin
        dbText "Name" ="SellerDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyerzip"
        dbLong "AggregateType" ="-1"
        dbMemo "Caption" ="buyerzip"
    End
    Begin
        dbText "Name" ="Expr1024"
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
        dbText "Name" ="Rmkt.pbos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="7785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="7785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionMed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1023"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-116
    Top =96
    Right =1142
    Bottom =742
    Left =-1
    Top =-1
    Right =1240
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-9
        Top =-1
        Right =193
        Bottom =283
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =207
        Top =1
        Right =351
        Bottom =92
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =204
        Top =104
        Right =348
        Bottom =249
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =382
        Top =5
        Right =509
        Bottom =339
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =519
        Top =3
        Right =650
        Bottom =215
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =693
        Top =7
        Right =837
        Bottom =151
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =865
        Top =11
        Right =965
        Bottom =157
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =697
        Top =158
        Right =843
        Bottom =471
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =1013
        Top =12
        Right =1157
        Bottom =156
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
