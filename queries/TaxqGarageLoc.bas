Operation =1
Option =0
Where ="(((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"b\" Or (Units.unitstatus)=\"c"
    "\" Or (Units.unitstatus)=\"tr\"))"
Having ="(((Clients.ClientShNm) Not Like \"demo*\" And (Clients.ClientShNm)<>\"test\") AN"
    "D ((Clients.Clienttype)=\"client\" Or (Clients.Clienttype)=\"guarantor\") AND (("
    "Locations.locobsolete)=No))"
Begin InputTables
    Name ="Clients"
    Name ="Locations"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Locations.LState"
    Expression ="Locations.LCounty"
    Expression ="Locations.LCity"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LZip"
    Expression ="Clients.ClientShNm"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocationName"
    Expression ="Locations.LTaxRate"
    Alias ="CountOfUnitID"
    Expression ="Count(Units.UnitID)"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Units"
    Expression ="Locations.LocationID = Units.UnitGarageLocationRef"
    Flag =1
End
Begin OrderBy
    Expression ="Locations.LState"
    Flag =0
    Expression ="Locations.LCounty"
    Flag =0
    Expression ="Locations.LCity"
    Flag =0
    Expression ="Locations.LZip"
    Flag =0
End
Begin Groups
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LCounty"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LAddrLine1"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Locations.LocationType"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LTaxRate"
    GroupLevel =0
    Expression ="Clients.Clienttype"
    GroupLevel =0
    Expression ="Locations.locobsolete"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="-1"
dbMemo "Filter" ="([TaxqGarageLoc].[ClientShNm]=\"AIR PRODUCTS\")"
dbMemo "OrderBy" ="[TaxqGarageLoc].[ClientShNm]"
Begin
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbInteger "ColumnOrder" ="8"
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
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =437
        Bottom =245
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =485
        Top =12
        Right =629
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
