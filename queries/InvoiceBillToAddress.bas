Operation =1
Option =0
Where ="(((Locations.LocObsolete)=No))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Alias ="BillingToID"
    Expression ="Locations.LocationID"
    Alias ="Location"
    Expression ="[LocationName] & \" \" & [LCity] & \", \" & [LState]"
    Alias ="BillingAddress"
    Expression ="[LocationStreet] & Chr(13) & Chr(10) & IIf(IsNull([LAddrLine1]),Null,[Locations]"
        ".[LAddrLine1] & Chr(13) & Chr(10)) & IIf(IsNull([Locations].[LAddrLine2]),Null,["
        "Locations].[LAddrLine2] & Chr(13) & Chr(10)) & [Locations].[LCity] & \" \" & [Lo"
        "cations].[LState] & \" \" & [Locations].[LZip]"
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
dbInteger "RowHeight" ="315"
Begin
    Begin
        dbText "Name" ="BillingAddress"
        dbInteger "ColumnWidth" ="4395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingToID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =146
    Top =43
    Right =1306
    Bottom =972
    Left =-1
    Top =-1
    Right =1142
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =482
        Bottom =432
        Top =0
        Name ="Locations"
        Name =""
    End
End
