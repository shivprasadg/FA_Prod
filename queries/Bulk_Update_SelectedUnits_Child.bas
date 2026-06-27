Operation =1
Option =8
Where ="(((ChildParts.ChildVendorREF) Is Not Null))"
Begin InputTables
    Name ="Bulk_Update_WorkTable"
    Name ="Units"
    Name ="ChildParts"
    Name ="Invoices"
End
Begin OutputColumns
    Expression ="ChildParts.*"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitVIN"
    Alias ="Desc"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [Uni"
        "tType])"
    Expression ="Invoices.InvoiceDate"
    Alias ="UID"
    Expression ="[ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass]"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="Bulk_Update_WorkTable"
    RightTable ="ChildParts"
    Expression ="Bulk_Update_WorkTable.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="Invoices"
    Expression ="ChildParts.ChildInvoiceNum = Invoices.InvoiceNumber"
    Flag =2
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
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.REFENG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.COMPRESSS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.MICROPROC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ELECDISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEXCofa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoices.InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.PartAddedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ExcludeCostDocsReports"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ImpactsResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ImpactsOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ATFChildSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ImpactsLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.cIFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.cIFDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1360
    Bottom =872
    Left =-1
    Top =-1
    Right =1260
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =21
        Top =4
        Right =165
        Bottom =148
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
    Begin
        Left =511
        Top =30
        Right =655
        Bottom =174
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =234
        Top =10
        Right =428
        Bottom =195
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =734
        Top =13
        Right =878
        Bottom =157
        Top =0
        Name ="Invoices"
        Name =""
    End
End
