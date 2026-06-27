Operation =1
Option =0
Begin InputTables
    Name ="CustodianTitleTrust1"
End
Begin OutputColumns
    Expression ="CustodianTitleTrust1.[FA-Asset Num]"
    Expression ="CustodianTitleTrust1.VIN"
    Expression ="CustodianTitleTrust1.[File Type]"
    Expression ="CustodianTitleTrust1.[Portfolio Certificate No]"
    Expression ="CustodianTitleTrust1.[Lender or Holder Name]"
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
        dbText "Name" ="CustodianTitleTrust1.[FA-Asset Num]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust1.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust1.[File Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust1.[Portfolio Certificate No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianTitleTrust1.[Lender or Holder Name]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =572
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CustodianTitleTrust1"
        Name =""
    End
End
