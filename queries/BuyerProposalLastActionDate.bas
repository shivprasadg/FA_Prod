Operation =1
Option =0
Begin InputTables
    Name ="BuyerProposalAutoStatus"
End
Begin OutputColumns
    Expression ="BuyerProposalAutoStatus.BuyerID"
    Alias ="LastProposalDate"
    Expression ="Max(BuyerProposalAutoStatus.LastAction)"
End
Begin Groups
    Expression ="BuyerProposalAutoStatus.BuyerID"
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
        dbText "Name" ="BuyerProposalAutoStatus.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastProposalDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1192
    Bottom =1021
    Left =-1
    Top =-1
    Right =1160
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =160
        Top =142
        Right =562
        Bottom =467
        Top =0
        Name ="BuyerProposalAutoStatus"
        Name =""
    End
End
