Operation =1
Option =0
Begin InputTables
    Name ="ClientMasterLeasesTypeCount"
End
Begin OutputColumns
    Expression ="ClientMasterLeasesTypeCount.clientID"
    Expression ="ClientMasterLeasesTypeCount.Client"
    Alias ="Type"
    Expression ="Switch(Max([ClientMasterLeasesTypeCount].[FA])=1 And Max([ClientMasterLeasesType"
        "Count].[NotFA])=1,\"Both\",Max([ClientMasterLeasesTypeCount].[FA])=1,\"FA\",Max("
        "[ClientMasterLeasesTypeCount].[NotFA])=1,\"Port\")"
    Alias ="Units"
    Expression ="Sum(ClientMasterLeasesTypeCount.Units)"
    Expression ="ClientMasterLeasesTypeCount.ClientGroupId"
End
Begin OrderBy
    Expression ="ClientMasterLeasesTypeCount.Client"
    Flag =0
End
Begin Groups
    Expression ="ClientMasterLeasesTypeCount.clientID"
    GroupLevel =0
    Expression ="ClientMasterLeasesTypeCount.Client"
    GroupLevel =0
    Expression ="ClientMasterLeasesTypeCount.ClientGroupId"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([ClientMasterLeasesTypeCount_Summed].[ClientGroupId]=60)"
Begin
    Begin
        dbText "Name" ="ClientMasterLeasesTypeCount.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientMasterLeasesTypeCount.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientMasterLeasesTypeCount.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =70
    Top =42
    Right =1674
    Bottom =993
    Left =-1
    Top =-1
    Right =1586
    Bottom =582
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =65
        Top =82
        Right =461
        Bottom =406
        Top =0
        Name ="ClientMasterLeasesTypeCount"
        Name =""
    End
End
