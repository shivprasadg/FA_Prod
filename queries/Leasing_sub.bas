Operation =1
Option =0
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Alias ="Expr1"
    Expression ="Clients.LeaseID"
    Alias ="Expr2"
    Expression ="Clients.LeaseDate"
    Alias ="Expr3"
    Expression ="Clients.LeaseTaxID"
    Expression ="Clients.LeaseNameVer"
    Expression ="Clients.LeaseCGS"
    Expression ="Clients.LeaseSTofCorp"
    Expression ="Clients.LeaseEntity"
    Alias ="Expr4"
    Expression ="Clients.LeaseMLAContactREF"
    Alias ="Expr5"
    Expression ="Clients.LeaseMLASigRef"
    Alias ="Expr6"
    Expression ="Clients.LeaseMLASigTitle"
    Alias ="Expr7"
    Expression ="Clients.LeaseMLASent"
    Alias ="Expr8"
    Expression ="Clients.LeaseMLABack"
    Alias ="Expr9"
    Expression ="Clients.LeaseGuarantor"
    Alias ="Expr10"
    Expression ="Clients.LeaseGtySigREF"
    Alias ="Expr11"
    Expression ="Clients.LeaseGtySigTitle"
    Alias ="Expr12"
    Expression ="Clients.LeaseGtyTaxID"
    Alias ="Expr13"
    Expression ="Clients.LeaseGtyExp"
    Alias ="Expr14"
    Expression ="Clients.LeaseGtyDate"
    Alias ="Expr15"
    Expression ="Clients.LeaseGtyBasis"
    Alias ="Expr16"
    Expression ="Clients.LeaseGTYCap"
    Alias ="Expr17"
    Expression ="Clients.LeaseGTYUnlim"
    Alias ="Expr18"
    Expression ="Clients.LeaseGtySent"
    Alias ="Expr19"
    Expression ="Clients.LeaseGtyBack"
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
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr19"
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
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =312
        Bottom =271
        Top =0
        Name ="Clients"
        Name =""
    End
End
