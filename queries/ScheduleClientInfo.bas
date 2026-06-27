Operation =1
Option =0
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientWebsite"
    Expression ="Clients.clientIndustry"
    Expression ="Clients.clientType"
    Expression ="Clients.clientMainPhone"
    Expression ="Clients.clientAnnualSales"
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
        dbText "Name" ="Clients.clientCompanyName"
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
        dbText "Name" ="Clients.clientWebsite"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientIndustry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientMainPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAnnualSales"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1587
    Bottom =872
    Left =-1
    Top =-1
    Right =1555
    Bottom =532
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =40
        Top =64
        Right =444
        Bottom =526
        Top =0
        Name ="Clients"
        Name =""
    End
End
