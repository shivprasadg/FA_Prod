Operation =1
Option =0
Having ="(((Nz([sgrpfaid],[schfaid])) Is Not Null) AND ((MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Group_UnitCount_Cost_Rents"
End
Begin OutputColumns
    Alias ="DateGroupAdded"
    Expression ="SchGrp.GroupAddedDate"
    Alias ="OPSQtr"
    Expression ="DatePart('q',[GroupAddedDate])"
    Alias ="OPSYear"
    Expression ="DatePart('yyyy',[GroupAddedDate])"
    Alias ="LFAID"
    Expression ="Nz([sgrpfaid],[schfaid])"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="Group_UnitCount_Cost_Rents.UnitCount"
    Alias ="TotalLORent"
    Expression ="Group_UnitCount_Cost_Rents.LORent"
    Alias ="TotalSyndRent"
    Expression ="Group_UnitCount_Cost_Rents.SyndRent"
    Expression ="Group_UnitCount_Cost_Rents.UnitCost"
    Expression ="Group_UnitCount_Cost_Rents.ChildPartCost"
    Alias ="TotalUnitCost"
    Expression ="[UnitCost]+[ChildPartCost]"
    Expression ="MstrLease.MLID"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
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
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Group_UnitCount_Cost_Rents"
    Expression ="SchGrp.SGrpID = Group_UnitCount_Cost_Rents.GroupID"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.GroupAddedDate"
    GroupLevel =0
    Expression ="DatePart('q',[GroupAddedDate])"
    GroupLevel =0
    Expression ="DatePart('yyyy',[GroupAddedDate])"
    GroupLevel =0
    Expression ="Nz([sgrpfaid],[schfaid])"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="Group_UnitCount_Cost_Rents.UnitCount"
    GroupLevel =0
    Expression ="Group_UnitCount_Cost_Rents.LORent"
    GroupLevel =0
    Expression ="Group_UnitCount_Cost_Rents.SyndRent"
    GroupLevel =0
    Expression ="Group_UnitCount_Cost_Rents.UnitCost"
    GroupLevel =0
    Expression ="Group_UnitCount_Cost_Rents.ChildPartCost"
    GroupLevel =0
    Expression ="[UnitCost]+[ChildPartCost]"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
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
        dbText "Name" ="LFAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSyndRent"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group_UnitCount_Cost_Rents.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalLORent"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalUnitCost"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group_UnitCount_Cost_Rents.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateGroupAdded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OPSYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group_UnitCount_Cost_Rents.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OPSQtr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1486
    Bottom =1038
    Left =-1
    Top =-1
    Right =1454
    Bottom =153
    Left =0
    Top =0
    ColumnsShown =543
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
        Bottom =219
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =831
        Top =69
        Right =975
        Bottom =213
        Top =0
        Name ="Group_UnitCount_Cost_Rents"
        Name =""
    End
End
