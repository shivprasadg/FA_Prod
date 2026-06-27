Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((Deals.dealOutcome)=\"Awarded\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Deals"
End
Begin OutputColumns
    Alias ="Lease No"
    Expression ="vw_SixKeys.MLNo"
    Alias ="Lessee"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="MstrLeaseDate"
    Expression ="vw_SixKeys.MLDate"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.PurchaseOrderDate"
    Alias ="DealFAID"
    Expression ="Deals.FAID"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeDate"
    Expression ="Deals.dealRALBack"
    Expression ="Deals.dealRFPDate"
    Expression ="Deals.dealLPAFcomp"
    Expression ="Deals.dealLPAFDue"
    Expression ="Deals.dealLPAFSent"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Deals"
    Expression ="vw_SixKeys.FAID = Deals.FAID"
    Flag =2
End
Begin OrderBy
    Expression ="Deals.dealOutcomeDate"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.MLDate"
    GroupLevel =0
    Expression ="vw_SixKeys.FAID"
    GroupLevel =0
    Expression ="vw_SixKeys.PurchaseOrderDate"
    GroupLevel =0
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
    GroupLevel =0
    Expression ="Deals.dealOutcomeDate"
    GroupLevel =0
    Expression ="Deals.dealRALBack"
    GroupLevel =0
    Expression ="Deals.dealRFPDate"
    GroupLevel =0
    Expression ="Deals.dealLPAFcomp"
    GroupLevel =0
    Expression ="Deals.dealLPAFDue"
    GroupLevel =0
    Expression ="Deals.dealLPAFSent"
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
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="3300"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease No"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLeaseDate"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRALBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFcomp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFDue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealFAID"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =34
    Right =1722
    Bottom =750
    Left =-1
    Top =-1
    Right =1571
    Bottom =381
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =23
        Top =-5
        Right =167
        Bottom =399
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =224
        Top =24
        Right =512
        Bottom =406
        Top =0
        Name ="Deals"
        Name =""
    End
End
