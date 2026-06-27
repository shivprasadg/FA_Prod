Operation =1
Option =0
Where ="(((Units.UnitID)=[forms]![UnitsForm]![unitid]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="Units"
    Name ="LocChangeT"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.Unitvin"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="LocChangeT.OLocationStreet"
    Expression ="LocChangeT.OCity"
    Expression ="LocChangeT.OCounty"
    Expression ="LocChangeT.OSt"
    Expression ="LocChangeT.Ozip"
    Expression ="LocChangeT.NLocationStreet"
    Expression ="LocChangeT.NCity"
    Expression ="LocChangeT.NCounty"
    Expression ="LocChangeT.NSt"
    Expression ="LocChangeT.Nzip"
    Expression ="LocChangeT.LocChDt"
    Expression ="LocChangeT.LocRptDt"
    Expression ="LocChangeT.OldTxstatus"
    Expression ="LocChangeT.OTaxNotes"
    Expression ="LocChangeT.Newtaxstatus"
    Expression ="LocChangeT.NewTaxnotes"
    Alias ="Assignee Bank"
    Expression ="Clients_1.clientCompanyName"
    Expression ="LocChangeT.ID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Clients_1"
    Expression ="SchGrp.Assignee_Bank = Clients_1.ClientShNm"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="LocChangeT"
    Expression ="Units.UnitID = LocChangeT.unitref"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="LocChangeT.ID"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee Bank"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =982
    Bottom =872
    Left =-1
    Top =-1
    Right =874
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =48
        Top =200
        Right =192
        Bottom =344
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =318
        Top =217
        Right =462
        Bottom =361
        Top =0
        Name ="LocChangeT"
        Name =""
    End
End
