Operation =1
Option =0
Begin InputTables
    Name ="DrawableUnitInvoices0"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="DrawableUnitInvoices0.Vendor"
    Alias ="Expr2"
    Expression ="DrawableUnitInvoices0.Invoice"
    Alias ="Expr3"
    Expression ="DrawableUnitInvoices0.InvCost"
    Alias ="Avail"
    Expression ="[Units]-CDbl(Nz([Drawn],0))"
    Alias ="Drawn"
    Expression ="CDbl(Nz([DrawableUnitInvoices0].[Drawn],0))"
    Alias ="Expr4"
    Expression ="DrawableUnitInvoices0.Units"
    Alias ="Expr5"
    Expression ="DrawableUnitInvoices0.SchID"
End
Begin OrderBy
    Expression ="DrawableUnitInvoices0.Invoice"
    Flag =0
End
Begin Groups
    Expression ="DrawableUnitInvoices0.Vendor"
    GroupLevel =0
    Expression ="DrawableUnitInvoices0.Invoice"
    GroupLevel =0
    Expression ="DrawableUnitInvoices0.InvCost"
    GroupLevel =0
    Expression ="[Units]-CDbl(Nz([Drawn],0))"
    GroupLevel =0
    Expression ="CDbl(Nz([DrawableUnitInvoices0].[Drawn],0))"
    GroupLevel =0
    Expression ="DrawableUnitInvoices0.Units"
    GroupLevel =0
    Expression ="DrawableUnitInvoices0.SchID"
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
dbMemo "Filter" ="(DrawableUnitInvoices.Drawn<>0 Or DrawableUnitInvoices.Drawn Is Null)"
Begin
    Begin
        dbText "Name" ="DrawableUnitInvoices0.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawableUnitInvoices0.Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawableUnitInvoices0.InvCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawableUnitInvoices0.Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawableUnitInvoices0.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drawn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =749
    Top =234
    Right =1696
    Bottom =754
    Left =-1
    Top =-1
    Right =915
    Bottom =164
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =2
        Top =18
        Right =270
        Bottom =226
        Top =0
        Name ="DrawableUnitInvoices0"
        Name =""
    End
End
