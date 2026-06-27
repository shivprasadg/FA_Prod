Operation =1
Option =0
Begin InputTables
    Name ="PDiemBill_Part1"
    Name ="BillStEndq"
    Name ="BillUnitq"
End
Begin OutputColumns
    Expression ="PDiemBill_Part1.BillID"
    Expression ="PDiemBill_Part1.clientID"
    Expression ="PDiemBill_Part1.MLID"
    Expression ="PDiemBill_Part1.SchID"
    Expression ="PDiemBill_Part1.SGrpID"
    Expression ="PDiemBill_Part1.UnitID"
    Expression ="PDiemBill_Part1.MLNo"
    Expression ="PDiemBill_Part1.SchNo"
    Expression ="PDiemBill_Part1.SGrpNO"
    Expression ="PDiemBill_Part1.BLCD"
    Expression ="PDiemBill_Part1.SGrpNoSlv"
    Expression ="PDiemBill_Part1.SGRPNoPD"
    Expression ="PDiemBill_Part1.unitunitnum"
    Expression ="PDiemBill_Part1.UnitVIN"
    Expression ="PDiemBill_Part1.UnitStatus"
    Expression ="PDiemBill_Part1.UnitAcceptDate"
    Expression ="PDiemBill_Part1.Desc"
    Expression ="PDiemBill_Part1.UnitTaxExempt"
    Expression ="PDiemBill_Part1.clientCompanyName"
    Expression ="PDiemBill_Part1.FAVenorNo"
    Expression ="PDiemBill_Part1.SchBnkClNO"
    Expression ="PDiemBill_Part1.vin4"
    Expression ="PDiemBill_Part1.unitpdrent"
    Expression ="PDiemBill_Part1.SGrpFSLPmt"
    Expression ="PDiemBill_Part1.unitinserv"
    Expression ="PDiemBill_Part1.schcrexp"
    Expression ="PDiemBill_Part1.SchFSL"
    Expression ="PDiemBill_Part1.unitrent"
    Expression ="PDiemBill_Part1.unittax"
    Expression ="PDiemBill_Part1.UNITOFFLEASEDT"
    Expression ="PDiemBill_Part1.UnitPDRent"
    Expression ="PDiemBill_Part1.SchCurrRent"
    Expression ="PDiemBill_Part1.BillDueDt"
    Expression ="PDiemBill_Part1.DueonRec"
    Expression ="PDiemBill_Part1.Billdt"
    Expression ="PDiemBill_Part1.BillSent"
    Expression ="BillStEndq.Billstart"
    Expression ="BillStEndq.BillEndDt"
    Expression ="BillStEndq.SlvEnd"
    Expression ="BillStEndq.PDEND"
    Expression ="BillUnitq.UBStart"
    Alias ="SlvDays"
    Expression ="IIf([billstart]>[schcrexp],0,IIf([sgrpnoslv]=No,0,IIf([ubstart]>[slvend],0,IIf(["
        "ubstart]>[billenddt],0,DateDiff(\"d\",IIf([ubstart]>[billstart],[ubstart],[bills"
        "tart]),IIf([billenddt]<[slvend],[billenddt],[slvend])+1)))))"
    Alias ="UBSlvS"
    Expression ="IIf([ubstart]>[billstart],[ubstart],[billstart])"
    Alias ="ubslvend"
    Expression ="IIf([billenddt]<[slvend],[billenddt],[slvend])"
End
Begin Joins
    LeftTable ="PDiemBill_Part1"
    RightTable ="BillStEndq"
    Expression ="PDiemBill_Part1.BillID = BillStEndq.BillID"
    Flag =2
    LeftTable ="PDiemBill_Part1"
    RightTable ="BillUnitq"
    Expression ="PDiemBill_Part1.BillID = BillUnitq.BillID"
    Flag =2
End
Begin Groups
    Expression ="PDiemBill_Part1.BillID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.clientID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.MLID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SchID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SGrpID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitID"
    GroupLevel =0
    Expression ="PDiemBill_Part1.MLNo"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SchNo"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SGrpNO"
    GroupLevel =0
    Expression ="PDiemBill_Part1.BLCD"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SGrpNoSlv"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SGRPNoPD"
    GroupLevel =0
    Expression ="PDiemBill_Part1.unitunitnum"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitVIN"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitStatus"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitAcceptDate"
    GroupLevel =0
    Expression ="PDiemBill_Part1.Desc"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitTaxExempt"
    GroupLevel =0
    Expression ="PDiemBill_Part1.clientCompanyName"
    GroupLevel =0
    Expression ="PDiemBill_Part1.FAVenorNo"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SchBnkClNO"
    GroupLevel =0
    Expression ="PDiemBill_Part1.vin4"
    GroupLevel =0
    Expression ="PDiemBill_Part1.unitpdrent"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SGrpFSLPmt"
    GroupLevel =0
    Expression ="PDiemBill_Part1.unitinserv"
    GroupLevel =0
    Expression ="PDiemBill_Part1.schcrexp"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SchFSL"
    GroupLevel =0
    Expression ="PDiemBill_Part1.unitrent"
    GroupLevel =0
    Expression ="PDiemBill_Part1.unittax"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UNITOFFLEASEDT"
    GroupLevel =0
    Expression ="PDiemBill_Part1.UnitPDRent"
    GroupLevel =0
    Expression ="PDiemBill_Part1.SchCurrRent"
    GroupLevel =0
    Expression ="PDiemBill_Part1.BillDueDt"
    GroupLevel =0
    Expression ="PDiemBill_Part1.DueonRec"
    GroupLevel =0
    Expression ="PDiemBill_Part1.Billdt"
    GroupLevel =0
    Expression ="PDiemBill_Part1.BillSent"
    GroupLevel =0
    Expression ="BillStEndq.Billstart"
    GroupLevel =0
    Expression ="BillStEndq.BillEndDt"
    GroupLevel =0
    Expression ="BillStEndq.SlvEnd"
    GroupLevel =0
    Expression ="BillStEndq.PDEND"
    GroupLevel =0
    Expression ="BillUnitq.UBStart"
    GroupLevel =0
    Expression ="IIf([billstart]>[schcrexp],0,IIf([sgrpnoslv]=No,0,IIf([ubstart]>[slvend],0,IIf(["
        "ubstart]>[billenddt],0,DateDiff(\"d\",IIf([ubstart]>[billstart],[ubstart],[bills"
        "tart]),IIf([billenddt]<[slvend],[billenddt],[slvend])+1)))))"
    GroupLevel =0
    Expression ="IIf([ubstart]>[billstart],[ubstart],[billstart])"
    GroupLevel =0
    Expression ="IIf([billenddt]<[slvend],[billenddt],[slvend])"
    GroupLevel =0
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
        dbText "Name" ="BillStEndq.SlvEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.Billstart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.BillEndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.PDEND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillUnitq.UBStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SlvDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UBSlvS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ubslvend"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.BillSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGrpBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.FAVenorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unitpdrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unitinserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.schcrexp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unittax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.UNITOFFLEASEDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.unitcurrentpd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.SchCurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDiemBill_Part1.Billdt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =136
    Top =22
    Right =1746
    Bottom =854
    Left =-1
    Top =-1
    Right =1578
    Bottom =574
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =111
        Top =49
        Right =263
        Bottom =549
        Top =0
        Name ="PDiemBill_Part1"
        Name =""
    End
    Begin
        Left =305
        Top =51
        Right =426
        Bottom =267
        Top =0
        Name ="BillStEndq"
        Name =""
    End
    Begin
        Left =304
        Top =276
        Right =426
        Bottom =522
        Top =0
        Name ="BillUnitq"
        Name =""
    End
End
