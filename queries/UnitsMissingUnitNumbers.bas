Operation =1
Option =0
Where ="(((vw_SixKeys.Unitnum) Is Null Or (vw_SixKeys.Unitnum)=\"\") AND ((vw_SixKeys.St"
    "atus)<>\"B\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.UnitID"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Make"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =115
    Top =166
    Right =1586
    Bottom =993
    Left =-1
    Top =-1
    Right =1439
    Bottom =367
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =76
        Top =49
        Right =359
        Bottom =400
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
