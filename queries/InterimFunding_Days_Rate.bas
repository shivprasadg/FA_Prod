Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="UnitTotalOEC"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitStatus"
    Alias ="InterimAmount"
    Expression ="CCur(Nz([UnitIFAmt],0))"
    Expression ="Units.UnitIFAmt"
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="Units.GroupREF"
    Expression ="UnitTotalOEC.UnitTotalCost"
    Alias ="InterimIntRate"
    Expression ="CDbl(Nz([InterimLoanRate],0.025))"
    Alias ="InterimLoanDays"
    Expression ="IIf(IsDate([UnitIFDate]) And IsDate([FundDate]),DateDiff(\"d\",[UnitIFDate],[Fun"
        "dDate]),0)"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.LoanDate"
    Expression ="Units.UnitIFDate"
    Expression ="UnitTotalOEC.UnitCost"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Units"
    RightTable ="UnitTotalOEC"
    Expression ="Units.UnitID = UnitTotalOEC.AssetID"
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
        dbText "Name" ="Units.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimLoanDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="InterimIntRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitIFAmt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.[FundDate]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Loandate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =102
    Top =10
    Right =1182
    Bottom =938
    Left =-1
    Top =-1
    Right =1062
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =553
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =222
        Top =207
        Right =392
        Bottom =550
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =222
        Top =45
        Right =391
        Bottom =189
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
End
