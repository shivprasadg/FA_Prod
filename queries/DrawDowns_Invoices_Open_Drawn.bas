Operation =1
Option =0
Begin InputTables
    Name ="vw_DrawDowns_Invoices"
    Alias ="DDI"
    Name ="SELECT DISTINCTROW DrawCnt.UnitInvoiceNum, Count(*) AS Split FROM (SELECT DISTIN"
        "CT Units.UnitInvoiceNum, Units.DrawDownId FROM Units)  AS DrawCnt WHERE (((DrawC"
        "nt.DrawDownId)>0)) GROUP BY DrawCnt.UnitInvoiceNum HAVING (((Count(*))>1))"
    Alias ="IsSplit"
End
Begin OutputColumns
    Expression ="DDI.Invoice"
    Expression ="DDI.Vendor"
    Alias ="SchGrp"
    Expression ="[Schedule]+' | '+[UnitGroup]"
    Alias ="Open"
    Expression ="Max(DDI.Open)"
    Alias ="Open Amt"
    Expression ="Format(Sum(IIf(IsNull([DDI].[DrawDownId]),[VendorCost],0)),'$#,###')"
    Alias ="Drawn"
    Expression ="Max(DDI.Drawn)"
    Alias ="Drawn Amt"
    Expression ="Format(Sum(IIf(IsNull([DDI].[DrawDownId]),0,[IFAmount])),'$#,###')"
    Alias ="DrawDownId"
    Expression ="DDI.DrawDownId"
    Expression ="DDI.VendorID"
    Expression ="DDI.ClientID"
    Expression ="DDI.MLID"
    Expression ="DDI.SchID"
    Expression ="DDI.sGrpID"
    Expression ="IsSplit.Split"
End
Begin Joins
    LeftTable ="DDI"
    RightTable ="IsSplit"
    Expression ="DDI.Invoice = IsSplit.UnitInvoiceNum"
    Flag =2
End
Begin Groups
    Expression ="DDI.Invoice"
    GroupLevel =0
    Expression ="DDI.Vendor"
    GroupLevel =0
    Expression ="[Schedule]+' | '+[UnitGroup]"
    GroupLevel =0
    Expression ="DDI.DrawDownId"
    GroupLevel =0
    Expression ="DDI.VendorID"
    GroupLevel =0
    Expression ="DDI.ClientID"
    GroupLevel =0
    Expression ="DDI.MLID"
    GroupLevel =0
    Expression ="DDI.SchID"
    GroupLevel =0
    Expression ="DDI.sGrpID"
    GroupLevel =0
    Expression ="IsSplit.Split"
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
        dbText "Name" ="DDI.Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Open"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDI.sGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDI.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Open Amt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drawn Amt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DDI.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDI.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDI.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drawn"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDI.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IsSplit.Split"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =16
    Top =190
    Right =1587
    Bottom =1080
    Left =-1
    Top =-1
    Right =1547
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =464
        Top =0
        Name ="DDI"
        Name =""
    End
    Begin
        Left =231
        Top =13
        Right =375
        Bottom =157
        Top =0
        Name ="IsSplit"
        Name =""
    End
End
