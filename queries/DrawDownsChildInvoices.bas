Operation =1
Option =0
Begin InputTables
    Name ="DrawDownsChildInvoices0"
End
Begin OutputColumns
    Expression ="DrawDownsChildInvoices0.Invoice"
    Expression ="DrawDownsChildInvoices0.Vendor"
    Expression ="DrawDownsChildInvoices0.InvCost"
    Alias ="Avail"
    Expression ="[Parts]-CDbl(Nz([Drawn],0))"
    Alias ="Drawn"
    Expression ="CDbl(Nz([DrawDownsChildInvoices0].[Drawn],0))"
    Expression ="DrawDownsChildInvoices0.Parts"
    Expression ="DrawDownsChildInvoices0.SchID"
    Alias ="GroupID"
    Expression ="DrawDownsChildInvoices0.SGrpID"
End
Begin OrderBy
    Expression ="DrawDownsChildInvoices0.Invoice"
    Flag =0
End
Begin Groups
    Expression ="DrawDownsChildInvoices0.Invoice"
    GroupLevel =0
    Expression ="DrawDownsChildInvoices0.Vendor"
    GroupLevel =0
    Expression ="DrawDownsChildInvoices0.InvCost"
    GroupLevel =0
    Expression ="[Parts]-CDbl(Nz([Drawn],0))"
    GroupLevel =0
    Expression ="CDbl(Nz([DrawDownsChildInvoices0].[Drawn],0))"
    GroupLevel =0
    Expression ="DrawDownsChildInvoices0.Parts"
    GroupLevel =0
    Expression ="DrawDownsChildInvoices0.SchID"
    GroupLevel =0
    Expression ="DrawDownsChildInvoices0.SGrpID"
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
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drawn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownsChildInvoices0.InvCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownsChildInvoices0.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownsChildInvoices0.Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownsChildInvoices0.Parts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownsChildInvoices0.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =170
    Top =171
    Right =1584
    Bottom =940
    Left =-1
    Top =-1
    Right =1390
    Bottom =149
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =375
        Top =43
        Right =519
        Bottom =187
        Top =0
        Name ="DrawDownsChildInvoices0"
        Name =""
    End
End
