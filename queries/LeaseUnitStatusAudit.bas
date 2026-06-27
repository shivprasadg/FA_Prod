Operation =1
Option =0
Where ="(((vw_SixKeys.Status) Not Like \"*K*\" And (vw_SixKeys.Status) Not Like \"*R*\")"
    " AND ((Units.UnitCOABack) Is Not Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.MYear"
    Alias ="Desc"
    Expression ="[Make] & \" - \" & [Model] & \" - \" & [type]"
    Expression ="Units.UnitCOABack"
    Alias ="Error"
    Expression ="Switch(IsDate([UnitCOABack]) And [Status]=\"B\",\"Should be 'A'\")"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.FAID"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4860"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.unitcoaback"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Error"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-43
    Top =-13
    Right =1476
    Bottom =993
    Left =-1
    Top =-1
    Right =1487
    Bottom =475
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =34
        Top =14
        Right =225
        Bottom =468
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =273
        Top =133
        Right =417
        Bottom =277
        Top =0
        Name ="Units"
        Name =""
    End
End
