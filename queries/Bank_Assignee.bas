Operation =1
Option =0
Where ="(((Clients.clientType)=\"bank\") AND ((Locations.LocationType)=\"Billing\") AND "
    "((Locations.LocObsolete)=False))"
Begin InputTables
    Name ="Clients"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Alias ="ShortName"
    Expression ="Clients.ClientShNm"
    Alias ="BankName"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientType"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocationStreet"
    Alias ="BankAddress"
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & [LAddrLine1] & Chr(13) & Chr(10) & IIf"
        "(IsNull([LAddrLine2]),Null,[LAddrLine2] & Chr(13) & Chr(10)) & IIf(IsNull([LAddr"
        "Line3]),Null,[LAddrLine3] & Chr(13) & Chr(10)) & [LCity] & \" \" & [LState] & \""
        " \" & Left([LZip],5) & IIf(Len([LZip])>5,\"-\" & Right([LZip],4),Null)"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LAddrLine3"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Locations.LocObsolete"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
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
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankAddress"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =286
    Top =25
    Right =1466
    Bottom =606
    Left =-1
    Top =-1
    Right =1162
    Bottom =247
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =109
        Top =28
        Right =290
        Bottom =279
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =403
        Top =-1
        Right =818
        Bottom =221
        Top =0
        Name ="Locations"
        Name =""
    End
End
