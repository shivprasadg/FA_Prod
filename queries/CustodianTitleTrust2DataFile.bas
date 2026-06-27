Operation =1
Option =0
Begin InputTables
    Name ="CustodianTitleTrust2"
End
Begin OutputColumns
    Alias ="FA-Asset Num"
    Expression ="CustodianTitleTrust2.[FA TitleLog Num]"
    Expression ="CustodianTitleTrust2.VIN"
    Expression ="CustodianTitleTrust2.[File Type]"
    Expression ="CustodianTitleTrust2.[Portfolio Certificate No]"
    Expression ="CustodianTitleTrust2.[Lender or Holder Name]"
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
        dbText "Name" ="CustodianTitleTrust2.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust2.[File Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust2.[Portfolio Certificate No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust2.[Lender or Holder Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA-Asset Num"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1595
    Bottom =963
    Left =-1
    Top =-1
    Right =1577
    Bottom =589
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =251
        Bottom =439
        Top =0
        Name ="CustodianTitleTrust2"
        Name =""
    End
End
