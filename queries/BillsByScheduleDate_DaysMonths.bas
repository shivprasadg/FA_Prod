Operation =1
Option =0
Where ="(((BillsByScheduleDate_Base.BillType)<>'' Or (BillsByScheduleDate_Base.BillType)"
    " Is Not Null))"
Begin InputTables
    Name ="BillsByScheduleDate_Base"
End
Begin OutputColumns
    Expression ="BillsByScheduleDate_Base.BilledSchID"
    Alias ="PerdiemDays"
    Expression ="IIf([billType] Like '*Perdiem*' Or [billType] Like '*PD*',DateDiff('d',[BilledFr"
        "om],[Billedto])+1,0)"
    Alias ="PDDays2015"
    Expression ="IIf(([billType]='Perdiem' Or [billType] Like '*PD*') And DatePart('yyyy',[billed"
        "From])=2015,DateDiff('d',[BilledFrom],[Billedto]+1),0)"
    Alias ="PDDays2014"
    Expression ="IIf(([billType]='Perdiem' Or [billType] Like '*PD*') And DatePart('yyyy',[billed"
        "From])=2014,DateDiff('d',[BilledFrom],[Billedto]+1),0)"
    Alias ="PDDays2013"
    Expression ="IIf(([billType]='Perdiem' Or [billType] Like '*PD*') And DatePart('yyyy',[billed"
        "From])=2013,DateDiff('d',[BilledFrom],[Billedto]+1),0)"
    Alias ="RemitMths"
    Expression ="IIf([billType] Like '*Remit*' Or [billType] Like '*Warr*' Or [BILLTYPE]='AnnRent"
        "',DateDiff('m',[BilledFrom],[Billedto]),0)"
    Alias ="Remit2015"
    Expression ="IIf(([billType] Like '*remit*' Or [billType] Like '*Warr*' Or [billtype]='AnnRen"
        "t') And DatePart('yyyy',[billedFrom])=2015,DateDiff('m',[BilledFrom],[Billedto])"
        ",0)"
    Alias ="Remit2014"
    Expression ="IIf(([billType] Like '*remit*' Or [billType] Like '*Warr*' Or [billtype]='AnnRen"
        "t') And DatePart('yyyy',[billedFrom])=2014,DateDiff('m',[BilledFrom],[Billedto])"
        ",0)"
    Alias ="Remit2013"
    Expression ="IIf(([billType] Like '*remit*' Or [billType] Like '*Warr*' Or [billtype]='AnnRen"
        "t') And DatePart('yyyy',[billedFrom])=2013,DateDiff('m',[BilledFrom],[Billedto])"
        ",0)"
    Alias ="FeeOnly2015"
    Expression ="IIf(([billType] Like '*Fee*') And DatePart('yyyy',[billedFrom])=2015,DateDiff('d"
        "',[BilledFrom],[Billedto]),0)"
    Alias ="FeeOnly2014"
    Expression ="IIf(([billType] Like '*Fee*') And DatePart('yyyy',[billedFrom])=2014,DateDiff('d"
        "',[BilledFrom],[Billedto]),0)"
    Alias ="FeeOnly2013"
    Expression ="IIf(([billType] Like '*Fee*') And DatePart('yyyy',[billedFrom])=2013,DateDiff('d"
        "',[BilledFrom],[Billedto]),0)"
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
        dbText "Name" ="BillsByScheduleDate_Base.BilledSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDDays2015"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remit2015"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDDays2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDDays2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemitMths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remit2014"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remit2013"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeOnly2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeOnly2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeOnly2015"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =125
    Top =111
    Right =1823
    Bottom =630
    Left =-1
    Top =-1
    Right =1666
    Bottom =210
    Left =1442
    Top =0
    ColumnsShown =539
    Begin
        Left =56
        Top =37
        Right =203
        Bottom =181
        Top =0
        Name ="BillsByScheduleDate_Base"
        Name =""
    End
End
