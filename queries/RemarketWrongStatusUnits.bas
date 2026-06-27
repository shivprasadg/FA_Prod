Operation =1
Option =0
Where ="(((RemarketUnitStatusAudit.BuyerId) Is Not Null) AND ((RemarketUnitStatusAudit.S"
    "tatus) Not Like \"*S*\" And (RemarketUnitStatusAudit.Status) Not Like \"*R*\" An"
    "d (RemarketUnitStatusAudit.Status)<>\"D\") AND ((RemarketUnitStatusAudit.unitoff"
    "leasedt) Is Null)) OR (((RemarketUnitStatusAudit.InvoicePaidDate) Is Null) AND ("
    "(RemarketUnitStatusAudit.BuyerId) Is Not Null) AND ((RemarketUnitStatusAudit.Sta"
    "tus)=\"S\" Or (RemarketUnitStatusAudit.Status)=\"KS\" Or (RemarketUnitStatusAudi"
    "t.Status)=\"D\") AND ((RemarketUnitStatusAudit.unitoffleasedt) Is Null)) OR (((R"
    "emarketUnitStatusAudit.Status) Like \"*S*\" Or (RemarketUnitStatusAudit.Status)="
    "\"D\") AND ((RemarketUnitStatusAudit.unitoffleasedt) Is Null))"
Begin InputTables
    Name ="RemarketUnitStatusAudit"
End
Begin OutputColumns
    Expression ="RemarketUnitStatusAudit.*"
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
        dbText "Name" ="RemarketUnitStatusAudit.Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Desc"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.BuyerId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Rmkt.ParkDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketUnitStatusAudit.Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =314
    Top =-9
    Right =1909
    Bottom =823
    Left =-1
    Top =-1
    Right =1571
    Bottom =447
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =129
        Top =79
        Right =549
        Bottom =442
        Top =0
        Name ="RemarketUnitStatusAudit"
        Name =""
    End
End
