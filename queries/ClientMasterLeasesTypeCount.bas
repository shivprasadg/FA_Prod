Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.MLNo"
    Alias ="FA"
    Expression ="IIf([MLOrig]=\"FA\",1,0)"
    Alias ="NotFA"
    Expression ="IIf([MLOrig]<>\"FA\",1,0)"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitId)"
    Expression ="vw_SixKeys.ClientGroupId"
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.MLOrig"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.MLID"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientGroupId"
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
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NotFA"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1604
    Bottom =991
    Left =-1
    Top =-1
    Right =1586
    Bottom =616
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =56
        Top =44
        Right =395
        Bottom =316
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
