Operation =1
Option =0
Where ="(((Clients.ClientGroupID)<> 18) AND ((Clients.clientType)=\"client\" Or (Clients"
    ".clientType) Like \"*target\"))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientIndustry"
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientType"
    Expression ="Clients.creditFAProfile"
    Expression ="Clients.creditFAProfile"
    Expression ="Clients.creditFAScore"
    Expression ="Clients.creditFAScore"
    Expression ="Clients.cpadded"
End
Begin OrderBy
    Expression ="Clients.clientIndustry"
    Flag =0
    Expression ="Clients.ClientGroupID"
    Flag =0
    Expression ="Clients.clientCompanyName"
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
        dbText "Name" ="Clients.[Client Group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="4200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientIndustry"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAProfile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAScore"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.cpadded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
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
    Bottom =289
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =334
        Bottom =404
        Top =0
        Name ="Clients"
        Name =""
    End
End
