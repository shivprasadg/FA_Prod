Operation =1
Option =0
Where ="(((Schedule.SchIn) Is Not Null And (Schedule.SchIn)>=[enter start dt] And (Sched"
    "ule.SchIn)<=[enter end date]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchOut"
    Expression ="Schedule.SchIn"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.SaleDate"
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
End
Begin OrderBy
    Expression ="Schedule.SchOut"
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
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
    Bottom =284
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =36
        Top =10
        Right =180
        Bottom =154
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =229
        Top =9
        Right =373
        Bottom =153
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =407
        Top =13
        Right =551
        Bottom =157
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =599
        Top =12
        Right =743
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
