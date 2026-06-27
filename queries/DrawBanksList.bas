Operation =1
Option =0
Begin InputTables
    Name ="Banks"
    Name ="DrawDowns"
End
Begin OutputColumns
    Expression ="Banks.BankId"
    Alias ="Bank"
    Expression ="Banks.BankShortName"
    Alias ="Draws"
    Expression ="IIf(Count([drawdowns].[DrawdownId])>0,\"Yes\",\"No\")"
End
Begin Joins
    LeftTable ="Banks"
    RightTable ="DrawDowns"
    Expression ="Banks.BankId = DrawDowns.DrawBankId"
    Flag =2
End
Begin OrderBy
    Expression ="Banks.BankShortName"
    Flag =0
End
Begin Groups
    Expression ="Banks.BankId"
    GroupLevel =0
    Expression ="Banks.BankShortName"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([DrawBanksList].[Draws]=\"Yes\")"
dbMemo "OrderBy" ="[DrawBanksList].[Bank]"
Begin
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Draws"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =63
    Top =52
    Right =1301
    Bottom =1041
    Left =-1
    Top =-1
    Right =1214
    Bottom =621
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =262
        Bottom =345
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =345
        Top =126
        Right =576
        Bottom =478
        Top =0
        Name ="DrawDowns"
        Name =""
    End
End
