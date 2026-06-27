Operation =1
Option =8
Where ="(((ChildParts.ChildVendorREF) Is Not Null))"
Begin InputTables
    Name ="Units"
    Name ="ChildParts"
    Name ="Invoices"
    Name ="Bulk_Update_WorkTable"
End
Begin OutputColumns
    Expression ="ChildParts.*"
    Expression ="Units.unitunitnum"
    Expression ="Units.unitvin"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
    Expression ="Invoices.InvoiceDate"
    Alias ="UID"
    Expression ="[ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass]"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="Invoices"
    Expression ="ChildParts.ChildInvRef = Invoices.Invoiceid"
    Flag =2
    LeftTable ="Units"
    RightTable ="Bulk_Update_WorkTable"
    Expression ="Units.UnitID = Bulk_Update_WorkTable.UnitID"
    Flag =1
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
        dbText "Name" ="ChildParts.ChildModel"
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
        dbText "Name" ="ChildParts.ChildMakeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModelRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVINLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEstDelDate"
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
        dbText "Name" ="ChildParts.ChildPDStartDate"
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
        dbText "Name" ="ChildParts.ARB ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildCofABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEXCofa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChDtadd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.CarbCompDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ExclFrGrpOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildDepInvDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildDepInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildDepPmtDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childadddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEstBld"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildActBld"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildFAPMT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildDrDon"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ReeferEngineModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ReeferEngineManufacturer"
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
        dbText "Name" ="unitunitnum"
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
    Right =1252
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =227
        Top =31
        Right =371
        Bottom =175
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =414
        Top =12
        Right =608
        Bottom =310
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =835
        Top =53
        Right =979
        Bottom =197
        Top =0
        Name ="Invoices"
        Name =""
    End
    Begin
        Left =1
        Top =-8
        Right =145
        Bottom =136
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
End
