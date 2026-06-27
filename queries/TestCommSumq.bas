Operation =1
Option =0
Where ="(((MstrLease.MLOrig)=\"fa\"))"
Having ="(((Clients.ClientShNm) Not Like \"demo*\"))"
Begin InputTables
    Name ="Salesperson"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="GrpUnitTypeq"
    Name ="SGrpExpq"
End
Begin OutputColumns
    Expression ="Salesperson.SPInitials"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchShortDesc"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Alias ="Expr1"
    Expression ="GrpUnitTypeq.CountOfUnitID"
    Expression ="GrpUnitTypeq.UnitCost"
    Alias ="Expr2"
    Expression ="GrpUnitTypeq.AvgOfUnitSynResid"
    Expression ="GrpUnitTypeq.AvgOfUNITRENT"
    Alias ="SumOfSGExp"
    Expression ="Sum(SGrpExpq.SGExp)"
End
Begin Joins
    LeftTable ="Salesperson"
    RightTable ="Clients"
    Expression ="Salesperson.SalespersonID = Clients.SalespersonREF"
    Flag =1
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="GrpUnitTypeq"
    Expression ="SchGrp.SGrpID = GrpUnitTypeq.UTGrpRef"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="SGrpExpq"
    Expression ="SchGrp.SGrpID = SGrpExpq.SGrpID"
    Flag =2
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
End
Begin Groups
    Expression ="Salesperson.SPInitials"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.SchShortDesc"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="GrpUnitTypeq.CountOfUnitID"
    GroupLevel =0
    Expression ="GrpUnitTypeq.UnitCost"
    GroupLevel =0
    Expression ="GrpUnitTypeq.AvgOfUnitSynResid"
    GroupLevel =0
    Expression ="GrpUnitTypeq.AvgOfUNITRENT"
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1530"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpUnitTypeq.UnitCost"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpUnitTypeq.AvgOfUNITRENT"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSGExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
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
    Bottom =390
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Salesperson"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1192
        Bottom =190
        Top =0
        Name ="GrpUnitTypeq"
        Name =""
    End
    Begin
        Left =1007
        Top =271
        Right =1151
        Bottom =415
        Top =0
        Name ="SGrpExpq"
        Name =""
    End
End
