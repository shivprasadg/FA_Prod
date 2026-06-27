Operation =1
Option =0
Where ="(((Units.UnitPlateNum) Is Null) AND ((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.St"
    "atus)=\"C\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLNo"
    Alias ="Assignee"
    Expression ="vw_SixKeys.BankShortName"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.Schedule"
    Alias ="Group"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Expression ="Units.UnitPlateNum"
    Expression ="vw_SixKeys.AcceptedDate"
    Expression ="Units.UnitCOASent"
    Alias ="COAOrig/Copy"
    Expression ="Nz([UnitCOABack],[UnitCofaCopy])"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.UnitDocAccpt"
    Alias ="Desc"
    Expression ="LookUp_UnitDesc.DescriptionMed"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="LookUp_UnitDesc"
    Expression ="Units.UnitID = LookUp_UnitDesc.AssetID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.AssetID"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="3795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofaCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COAOrig/Copy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionMed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =208
    Top =79
    Right =1988
    Bottom =836
    Left =-1
    Top =-1
    Right =1756
    Bottom =304
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =290
        Top =42
        Right =461
        Bottom =313
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =589
        Top =72
        Right =733
        Bottom =255
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
