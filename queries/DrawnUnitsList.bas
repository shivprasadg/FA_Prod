Operation =1
Option =8
Where ="(((DrawDowns.DrawDownId)=[Forms]![DrawDownManager]![DrawDownId]))"
Begin InputTables
    Name ="DrawDowns"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.AssetId"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Unitnum"
    Alias ="SchNum"
    Expression ="vw_SixKeys.Schedule"
    Alias ="Invoice"
    Expression ="vw_SixKeys.InvoiceNumber"
    Alias ="Cost"
    Expression ="vw_SixKeys.UnitCost"
    Expression ="vw_SixKeys.SchId"
    Expression ="DrawDowns.DrawDownId"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="vw_SixKeys"
    Expression ="DrawDowns.DrawDownId = vw_SixKeys.DrawDownId"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.Unitnum"
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
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbInteger "ColumnWidth" ="3000"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =175
    Top =72
    Right =1478
    Bottom =841
    Left =-1
    Top =-1
    Right =1271
    Bottom =384
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =44
        Top =39
        Right =188
        Bottom =183
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =597
        Top =103
        Right =741
        Bottom =402
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
