Operation =1
Option =0
Where ="(((Rmkt.BuyerRef)=forms!ReMarketInvoiceManager!buyerref) And ((Rmkt.InvoicePaidD"
    "ate) Is Not Null) And ((Rmkt.InvoiceRequestDate)=forms!ReMarketInvoiceManager!Re"
    "questDate) And ((Rmkt.pbos)=True))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
    Name ="Locations"
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
    Alias ="Buyerzip"
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
    Expression ="Rmkt.pbos"
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
    Flag =2
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
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LOCCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyerzip"
        dbMemo "Caption" ="buyerzip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.pbos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleBody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxOption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.PortalLocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleSubject"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleAttachmentComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingMethod"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleAttachmentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Longitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Latitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.SalesForceID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCityID"
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
    Right =1240
    Bottom =342
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =202
        Top =1
        Right =411
        Bottom =182
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =464
        Top =10
        Right =608
        Bottom =154
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =660
        Top =17
        Right =804
        Bottom =160
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =516
        Top =190
        Right =660
        Bottom =334
        Top =0
        Name ="Locations"
        Name =""
    End
End
