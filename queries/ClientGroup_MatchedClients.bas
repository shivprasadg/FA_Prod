Operation =1
Option =0
Begin InputTables
    Name ="vw_ClientGroup_MatchedClients"
End
Begin OutputColumns
    Expression ="vw_ClientGroup_MatchedClients.*"
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
        dbText "Name" ="vw_ClientGroup_MatchedClients.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ClientGroup_MatchedClients.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ClientGroup_MatchedClients.ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ClientGroup_MatchedClients.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ClientGroup_MatchedClients.ParentGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ClientGroup_MatchedClients.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =202
    Top =82
    Right =1721
    Bottom =1080
    Left =-1
    Top =-1
    Right =1487
    Bottom =664
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =223
        Top =102
        Right =367
        Bottom =313
        Top =0
        Name ="vw_ClientGroup_MatchedClients"
        Name =""
    End
End
