Operation =1
Option =1
Where ="(((Locations.LocationType)=\"Garage\"))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
End
Begin OrderBy
    Expression ="Locations.ClientREF"
    Flag =0
    Expression ="Locations.LocationName"
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
        dbText "Name" ="LocationAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
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
        dbText "Name" ="Locations.LCityID"
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
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.oldlocid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxOption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleBody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Latitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.District"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsCmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationBillingRuleSubject"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocBKACCT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.RLocId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LOCCCC"
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
        dbText "Name" ="Locations.LocSubLessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.PortalLocationName"
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
        dbText "Name" ="Locations.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.SalesForceID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-267
    Top =85
    Right =1158
    Bottom =1074
    Left =-1
    Top =-1
    Right =1407
    Bottom =587
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =266
        Top =-18
        Right =654
        Bottom =584
        Top =0
        Name ="Locations"
        Name =""
    End
End
