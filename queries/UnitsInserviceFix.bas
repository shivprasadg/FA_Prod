Operation =1
Option =0
Where ="(((Units.UnitAcceptDate) Is Not Null) AND ((Units.UnitPDStartDate) Is Null) AND "
    "((Units.UnitInServ) Is Null) AND ((Units.UnitStatus)=\"A\")) OR (((Units.UnitAcc"
    "eptDate) Is Null) AND ((Units.UnitPDStartDate) Is Not Null) AND ((Units.UnitInSe"
    "rv) Is Null) AND ((Units.UnitStatus)=\"A\")) OR (((Units.UnitAcceptDate) Is Not "
    "Null) AND ((Units.UnitPDStartDate) Is Not Null) AND ((Units.UnitInServ) Is Null)"
    " AND ((Units.UnitStatus)=\"A\"))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitPDStartDate"
    Expression ="Units.UnitInServ"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitInServInd"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1192
    Bottom =1021
    Left =-1
    Top =-1
    Right =1072
    Bottom =557
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =109
        Top =56
        Right =569
        Bottom =481
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =608
        Top =125
        Right =864
        Bottom =393
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
