Operation =1
Option =0
Having ="(((UnitDeliveredCost.GroupId) Is Not Null))"
Begin InputTables
    Name ="Clients_Sch_Group_OEC_Part1"
    Name ="UnitDeliveredCost"
    Name ="SchGrp"
    Name ="ScheduleGroupPOTypes"
End
Begin OutputColumns
    Expression ="Clients_Sch_Group_OEC_Part1.CompanyName"
    Expression ="Clients_Sch_Group_OEC_Part1.MLNo"
    Expression ="Clients_Sch_Group_OEC_Part1.Schedule"
    Expression ="Clients_Sch_Group_OEC_Part1.UnitGroup"
    Expression ="ScheduleGroupPOTypes.POType"
    Expression ="ScheduleGroupPOTypes.Ordered"
    Expression ="ScheduleGroupPOTypes.OrderedDate"
    Expression ="Clients_Sch_Group_OEC_Part1.Units"
    Alias ="UnitCost"
    Expression ="Sum(UnitDeliveredCost.UnitCost)"
    Alias ="ChildCost"
    Expression ="Sum(UnitDeliveredCost.ChildCost)"
    Alias ="DeliveredCost"
    Expression ="Sum(UnitDeliveredCost.DeliveredCost)"
    Alias ="SyndRent"
    Expression ="Sum(UnitDeliveredCost.SyndRent)"
    Alias ="LRF"
    Expression ="IIf([SyndRent]=0,0,Round([SyndRent]/[DeliveredCost],6))"
    Expression ="UnitDeliveredCost.GroupId"
End
Begin Joins
    LeftTable ="Clients_Sch_Group_OEC_Part1"
    RightTable ="UnitDeliveredCost"
    Expression ="Clients_Sch_Group_OEC_Part1.SGrpID = UnitDeliveredCost.GroupId"
    Flag =1
    LeftTable ="Clients_Sch_Group_OEC_Part1"
    RightTable ="SchGrp"
    Expression ="Clients_Sch_Group_OEC_Part1.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Clients_Sch_Group_OEC_Part1"
    RightTable ="ScheduleGroupPOTypes"
    Expression ="Clients_Sch_Group_OEC_Part1.SGrpID = ScheduleGroupPOTypes.GroupID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients_Sch_Group_OEC_Part1.CompanyName"
    Flag =0
    Expression ="Clients_Sch_Group_OEC_Part1.SchID"
    Flag =0
End
Begin Groups
    Expression ="Clients_Sch_Group_OEC_Part1.CompanyName"
    GroupLevel =0
    Expression ="Clients_Sch_Group_OEC_Part1.MLNo"
    GroupLevel =0
    Expression ="Clients_Sch_Group_OEC_Part1.Schedule"
    GroupLevel =0
    Expression ="Clients_Sch_Group_OEC_Part1.UnitGroup"
    GroupLevel =0
    Expression ="ScheduleGroupPOTypes.POType"
    GroupLevel =0
    Expression ="ScheduleGroupPOTypes.Ordered"
    GroupLevel =0
    Expression ="ScheduleGroupPOTypes.OrderedDate"
    GroupLevel =0
    Expression ="Clients_Sch_Group_OEC_Part1.Units"
    GroupLevel =0
    Expression ="UnitDeliveredCost.GroupId"
    GroupLevel =0
    Expression ="Clients_Sch_Group_OEC_Part1.SchID"
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
        dbText "Name" ="Clients_Sch_Group_OEC_Part1.Units"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="Clients_Sch_Group_OEC_Part1.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveredCost"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients_Sch_Group_OEC_Part1.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_Sch_Group_OEC_Part1.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_Sch_Group_OEC_Part1.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ChildCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDeliveredCost.GroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleGroupPOTypes.POType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleGroupPOTypes.Ordered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleGroupPOTypes.OrderedDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =158
    Top =83
    Right =1760
    Bottom =966
    Left =-1
    Top =-1
    Right =1570
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =103
        Top =94
        Right =247
        Bottom =356
        Top =0
        Name ="Clients_Sch_Group_OEC_Part1"
        Name =""
    End
    Begin
        Left =290
        Top =160
        Right =434
        Bottom =400
        Top =0
        Name ="UnitDeliveredCost"
        Name =""
    End
    Begin
        Left =471
        Top =263
        Right =615
        Bottom =521
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =743
        Top =244
        Right =887
        Bottom =388
        Top =0
        Name ="ScheduleGroupPOTypes"
        Name =""
    End
End
