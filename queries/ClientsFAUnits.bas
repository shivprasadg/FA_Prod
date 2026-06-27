Operation =1
Option =0
Where ="(((vw_SixKeys.Type)=\"Tractor\" Or (vw_SixKeys.Type)=\"Straight Truck\" Or (vw_S"
    "ixKeys.Type)=\"Van\"))"
Having ="(((vw_SixKeys.MLOrig)=\"fa\") AND ((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.Stat"
    "us)=\"TR\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="ClientGroups"
End
Begin OutputColumns
    Expression ="ClientGroups.ClientGroupName"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="FAUnits"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.clientID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="ClientGroups"
    Expression ="vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
End
Begin Groups
    Expression ="ClientGroups.ClientGroupName"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.Status"
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
Begin
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAUnits"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="ClientGroups.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstDeliveryDate"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =79
    Top =26
    Right =931
    Bottom =986
    Left =-1
    Top =-1
    Right =834
    Bottom =624
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =247
        Bottom =597
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =321
        Top =87
        Right =465
        Bottom =231
        Top =0
        Name ="ClientGroups"
        Name =""
    End
End
