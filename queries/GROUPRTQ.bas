Operation =1
Option =0
Begin InputTables
    Name ="ClientGroups"
End
Begin OutputColumns
    Expression ="ClientGroups.ClientGroupId"
    Expression ="ClientGroups.ClientGroupName"
End
Begin OrderBy
    Expression ="ClientGroups.ClientGroupName"
    Flag =0
End
Begin Groups
    Expression ="ClientGroups.ClientGroupId"
    GroupLevel =0
    Expression ="ClientGroups.ClientGroupName"
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
        dbText "Name" ="ClientGroups.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroups.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =920
    Bottom =679
    Left =-1
    Top =-1
    Right =888
    Bottom =229
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =385
        Top =12
        Right =529
        Bottom =156
        Top =0
        Name ="ClientGroups"
        Name =""
    End
End
