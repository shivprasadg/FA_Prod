Operation =1
Option =0
Where ="(((Units.UnitStatus)<>\"T\" And (Units.UnitStatus)<>\"O\"))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Rmkt"
    Name ="Banks"
End
Begin OutputColumns
    Expression ="Schedule.SchID"
    Alias ="GroupId"
    Expression ="SchGrp.SGrpID"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="BuyOutDate"
    Expression ="Format([rmkt].[BuyOutDate],\"mm/dd/yy\")"
    Alias ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Alias ="VIN"
    Expression ="Units.UnitVIN"
    Alias ="S"
    Expression ="Units.UnitStatus"
    Alias ="Bank"
    Expression ="Banks.BankShortName"
    Alias ="Loan#"
    Expression ="Rmkt.IFLoanNumber"
    Alias ="Make"
    Expression ="Units.unitmake"
    Alias ="UnitGroup"
    Expression ="SchGrp.SGrpNO"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Rmkt"
    RightTable ="Banks"
    Expression ="Rmkt.IFBankID = Banks.BankId"
    Flag =2
End
Begin OrderBy
    Expression ="Units.UnitID"
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
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Loan#"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="SellerInv"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyDateEffective"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOutDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-22
    Top =202
    Right =1261
    Bottom =925
    Left =-1
    Top =-1
    Right =1259
    Bottom =196
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =15
        Top =25
        Right =165
        Bottom =371
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =203
        Top =82
        Right =347
        Bottom =438
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =393
        Top =95
        Right =537
        Bottom =509
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =590
        Top =66
        Right =734
        Bottom =286
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =767
        Top =156
        Right =911
        Bottom =300
        Top =0
        Name ="Banks"
        Name =""
    End
End
