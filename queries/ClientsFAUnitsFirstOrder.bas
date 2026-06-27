Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="ClientGroups"
End
Begin OutputColumns
    Expression ="ClientGroups.ClientGroupName"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="ActualDeliveryDate"
    Expression ="Min(IIf(Nz([vw_SixKeys.ActualDeliveryDate],\"1/1/2999\")>[BLCD],[BLCD],[vw_SixKe"
        "ys.ActualDeliveryDate]))"
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
        dbText "Name" ="ClientGroups.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =87
    Top =108
    Right =1270
    Bottom =1077
    Left =-1
    Top =-1
    Right =1165
    Bottom =522
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
