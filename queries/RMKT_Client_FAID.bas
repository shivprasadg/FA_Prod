Operation =1
Option =0
Begin InputTables
    Name ="vw_Sch_Grp_Units_Remarketing"
End
Begin OutputColumns
    Expression ="vw_Sch_Grp_Units_Remarketing.clientID"
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    Alias ="FAID"
    Expression ="Nz([vw_Sch_Grp_Units_Remarketing].[ClientShortName] & \"_RMKT\")"
End
Begin Groups
    Expression ="vw_Sch_Grp_Units_Remarketing.clientID"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    GroupLevel =0
    Expression ="Nz([vw_Sch_Grp_Units_Remarketing].[ClientShortName] & \"_RMKT\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1571
    Bottom =855
    Left =-1
    Top =-1
    Right =1539
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_Sch_Grp_Units_Remarketing"
        Name =""
    End
End
