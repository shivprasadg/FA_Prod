Operation =1
Option =0
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
    Alias ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Alias ="VIN"
    Expression ="Units.UnitVIN"
    Alias ="Status"
    Expression ="Units.UnitStatus"
    Alias ="Bank"
    Expression ="Banks.BankShortName"
    Alias ="Loan#"
    Expression ="Rmkt.IFLoanNumber"
    Alias ="Make"
    Expression ="Units.unitmake"
    Alias ="MYear"
    Expression ="Units.Unityr"
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
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loan#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =186
    Top =17
    Right =1469
    Bottom =740
    Left =-1
    Top =-1
    Right =1251
    Bottom =236
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
        Left =597
        Top =102
        Right =741
        Bottom =322
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
