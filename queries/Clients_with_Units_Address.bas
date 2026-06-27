Operation =1
Option =0
Begin InputTables
    Name ="Client_with_Units"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Client_with_Units.ClientGroupId"
    Expression ="Clients.clientID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientAddress1"
    Expression ="Clients.clientAddress2"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientZip"
    Expression ="Clients.clientcounty"
    Expression ="Clients.clientCountry"
End
Begin Joins
    LeftTable ="Client_with_Units"
    RightTable ="Clients"
    Expression ="Client_with_Units.clientID = Clients.clientID"
    Flag =1
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientcounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_with_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1252
    Bottom =808
    Left =-1
    Top =-1
    Right =1144
    Bottom =359
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =24
        Right =206
        Bottom =168
        Top =0
        Name ="Client_with_Units"
        Name =""
    End
    Begin
        Left =241
        Top =93
        Right =512
        Bottom =394
        Top =0
        Name ="Clients"
        Name =""
    End
End
