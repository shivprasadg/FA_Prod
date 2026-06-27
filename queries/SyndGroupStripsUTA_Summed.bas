Operation =1
Option =0
Begin InputTables
    Name ="SyndGroupStripsUTA"
End
Begin OutputColumns
    Expression ="SyndGroupStripsUTA.ScheduleID"
    Expression ="SyndGroupStripsUTA.GroupID"
    Alias ="GroupStrips"
    Expression ="Max(SyndGroupStripsUTA.StripsNum)"
    Alias ="GroupStripAmount"
    Expression ="Sum(CCur(Nz([SalesRent],0)*Nz([StripsNum],0)))"
    Alias ="UTAPer"
    Expression ="Max(SyndGroupStripsUTA.UTA)"
    Alias ="UTAGroupTotal"
    Expression ="Sum(SyndGroupStripsUTA.UTA)"
End
Begin Groups
    Expression ="SyndGroupStripsUTA.ScheduleID"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA.GroupID"
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
dbMemo "Filter" ="([SyndGroupStripsUTA_Summed].[GroupID]=1494)"
Begin
    Begin
        dbText "Name" ="SyndGroupStripsUTA.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTAPer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTAGroupTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupStrips"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupStripAmount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SyndGroupStripsUTA.ScheduleID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =107
    Top =26
    Right =1290
    Bottom =1024
    Left =-1
    Top =-1
    Right =1151
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =191
        Top =193
        Right =455
        Bottom =436
        Top =0
        Name ="SyndGroupStripsUTA"
        Name =""
    End
End
