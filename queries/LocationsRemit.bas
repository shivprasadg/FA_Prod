Operation =1
Option =0
Where ="(((Locations.LocObsolete)=No) AND ((Locations.LocationName) Like \"*Remit*\"))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.*"
    Expression ="Locations.LocationName"
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
        dbText "Name" ="Locations.LocationType"
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
        dbText "Name" ="Locations.LOCCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LPhone"
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
        dbText "Name" ="Locations.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleAttachmentComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleAttachmentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.PortalLocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCityID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingMethod"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleBody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LFax"
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
        dbText "Name" ="Locations.LocationBillingRuleSubject"
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
        dbText "Name" ="Locations.SalesForceID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountyID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1365
    Bottom =728
    Left =-1
    Top =-1
    Right =1347
    Bottom =337
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =160
        Top =10
        Right =474
        Bottom =319
        Top =0
        Name ="Locations"
        Name =""
    End
End
