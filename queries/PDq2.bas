Operation =1
Option =0
Begin InputTables
    Name ="PDQ1a"
End
Begin OutputColumns
    Alias ="LineTot"
    Expression ="([PDDAYS]*[PDRENT]+IIf(IsNull([SLVRENT]),0,[slvrent]))*(1+IIf([SCHFSL]=True Or ["
        "UNITTAXEXEMPT]=True Or [STTAXTYPE]<>\"RR\",0,[LOCATIONS].[LTaxRate]))"
    Expression ="PDQ1a.*"
    Expression ="PDQ1a.LCity"
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
        dbText "Name" ="LineTot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SchGrp.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.EndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLoST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLocStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLoST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLocCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLocSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BLocZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.blOCst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.unitpdrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BillStdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.unitinserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Schedule.schcrexp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.INTDaysQ.SlvDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.SlvRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.DSlvRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.PdRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Dailytax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BillUnitq.UnitPDST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Units.unittax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.BillStEndq.SlvEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.PDDAYS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bills.BillSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.INTDaysQ.UBSlvS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bill1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDQ1a.Bill2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =267
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="PDQ1a"
        Name =""
    End
End
