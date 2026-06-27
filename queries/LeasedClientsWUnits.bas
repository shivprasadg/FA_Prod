Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.Status)=\"B\" Or (vw_SixKeys.Status)="
    "\"O\" Or (vw_SixKeys.Status)=\"C\" Or (vw_SixKeys.Status)=\"TR\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.CompanyName"
    Alias ="Unit Count"
    Expression ="Count(vw_SixKeys.AssetID)"
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyName"
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
        dbText "Name" ="Unit Count"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =152
    Top =115
    Right =1402
    Bottom =883
    Left =-1
    Top =-1
    Right =1218
    Bottom =281
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =126
        Top =48
        Right =429
        Bottom =192
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
