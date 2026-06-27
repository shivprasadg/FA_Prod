Operation =1
Option =2
Having ="(((Clients.ClientGroupID)=forms![po selection]!ClientGroupID) And ((Deals.dealOu"
    "tcome) Is Null)) Or (((Clients.ClientGroupID)=forms![po selection]!ClientGroupID"
    ") And ((Deals.dealOutcome) Is Not Null And (Deals.dealOutcome)<>\"dead\" And (De"
    "als.dealOutcome)<>\"lost\"))"
Begin InputTables
    Name ="Deals"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.ClientGroupID"
    Expression ="Deals.FAID"
    Expression ="Clients.ClientShNm"
    Expression ="Deals.dealOutcome"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="Clients"
    Expression ="Deals.clientRef = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Deals.FAID"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientGroupID"
    GroupLevel =0
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
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
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
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
    Right =1226
    Bottom =368
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =528
        Bottom =214
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
