Operation =1
Option =0
Begin InputTables
    Name ="Banks"
End
Begin OutputColumns
    Expression ="Banks.BankShortName"
    Alias ="BankAddress"
    Expression ="[BankName] & Chr(13) & Chr(10) & [Address1] & Chr(13) & Chr(10) & IIf(IsNull([Ad"
        "dress2]),Null,[Banks].[Address2] & Chr(13) & Chr(10)) & [Banks].[City] & \" \" &"
        " [Banks].[state] & \" \" & Left([Banks].[Zip],5)"
    Expression ="Banks.BankId"
    Expression ="Banks.Phone"
    Expression ="Banks.PortfolioNumber"
    Expression ="Banks.Fax"
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
dbInteger "RowHeight" ="360"
Begin
    Begin
        dbText "Name" ="Banks.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="bankAddress"
        dbInteger "ColumnWidth" ="4395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankShortName"
        dbInteger "ColumnWidth" ="6225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.PortfolioNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.Fax"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =151
    Top =-1
    Right =1729
    Bottom =928
    Left =-1
    Top =-1
    Right =1560
    Bottom =510
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =86
        Top =41
        Right =230
        Bottom =319
        Top =0
        Name ="Banks"
        Name =""
    End
End
