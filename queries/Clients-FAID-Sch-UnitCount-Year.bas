Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.CompanyType"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Schedule"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="Year"
    Expression ="CDbl(20 & Trim(Left([FAID],2)))"
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.FAID"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.ClientGroupId"
    GroupLevel =0
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyName"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyType"
    GroupLevel =0
    Expression ="vw_SixKeys.FAID"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
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
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-13
    Top =70
    Right =1381
    Bottom =898
    Left =-1
    Top =-1
    Right =1362
    Bottom =460
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =47
        Top =40
        Right =416
        Bottom =420
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
