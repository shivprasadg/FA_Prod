Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.SchId"
    Alias ="DrawnUnits"
    Expression ="Count(vw_SixKeys.AssetId)"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="vw_SixKeys"
    Expression ="DrawDowns.DrawDownId = vw_SixKeys.DrawDownId"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.FAID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchId"
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
        dbText "Name" ="DrawnUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =196
    Top =108
    Right =1461
    Bottom =831
    Left =-1
    Top =-1
    Right =1233
    Bottom =389
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =15
        Top =13
        Right =159
        Bottom =157
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =235
        Top =40
        Right =379
        Bottom =324
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
