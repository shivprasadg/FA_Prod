Operation =1
Option =0
Begin InputTables
    Name ="Banks"
End
Begin OutputColumns
    Expression ="Banks.BankName"
    Expression ="Banks.BankShortName"
    Expression ="Banks.BankType"
    Expression ="Banks.BankAddress"
    Expression ="Banks.Address1"
    Expression ="Banks.Address2"
    Expression ="Banks.City"
    Expression ="Banks.State"
    Expression ="Banks.Zip"
    Expression ="Banks.PortfolioNumber"
    Expression ="Banks.PortfolioNumberInterm"
    Expression ="Banks.BankID"
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
        dbText "Name" ="Banks.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankAddress"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.PortfolioNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.PortfolioNumberInterm"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =343
    Top =59
    Right =1766
    Bottom =1293
    Left =-1
    Top =-1
    Right =1399
    Bottom =920
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =91
        Top =51
        Right =514
        Bottom =442
        Top =0
        Name ="Banks"
        Name =""
    End
End
