Operation =1
Option =0
Begin InputTables
    Name ="RemarketSourcetoSold"
End
Begin OutputColumns
    Expression ="RemarketSourcetoSold.AssetID"
    Expression ="RemarketSourcetoSold.clientID"
    Expression ="RemarketSourcetoSold.BuyerId"
    Expression ="RemarketSourcetoSold.SourceCompany"
    Expression ="RemarketSourcetoSold.Buyer"
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
        dbText "Name" ="RemarketSourcetoSold.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketSourcetoSold.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketSourcetoSold.BuyerId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketSourcetoSold.Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketSourcetoSold.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1587
    Bottom =872
    Left =-1
    Top =-1
    Right =1517
    Bottom =469
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =187
        Top =107
        Right =331
        Bottom =463
        Top =0
        Name ="RemarketSourcetoSold"
        Name =""
    End
End
