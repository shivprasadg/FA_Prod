Operation =1
Option =0
Where ="(((DrawDowns.DrawDownId) Is Not Null))"
Begin InputTables
    Name ="DrawDowns"
    Name ="Units"
    Name ="vw_SixKeys"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="DrawDowns.DrawDownId"
    Alias ="Vendor"
    Expression ="Vendors.VendorSNm"
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="RequestNumber"
    Expression ="DrawDowns.DrawRequestNumber"
    Expression ="DrawDowns.DrawName"
    Alias ="Draw"
    Expression ="DrawDowns.DrawNumber"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="VIN"
    Expression ="Units.UnitVIN"
    Alias ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Alias ="UnitDesc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [Un"
        "itSubType]) & IIf([UnitType]=\"Tractor\",\"\",\" \" & [UnitType])"
    Alias ="UnitCost"
    Expression ="Units.UnitFinalCost"
    Alias ="Invoice"
    Expression ="Units.UnitInvoiceNum"
    Alias ="InvoiceDate"
    Expression ="Units.UnitInvoiceDate"
    Expression ="vw_SixKeys.UnitDescDocVer"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="vw_SixKeys.VendorId"
    Expression ="DrawDowns.DrawRequestNumber"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
    LeftTable ="DrawDowns"
    RightTable ="vw_SixKeys"
    Expression ="DrawDowns.DrawDownId = vw_SixKeys.DrawDownId"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Vendors"
    Expression ="vw_SixKeys.VendorId = Vendors.VendorID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="DrawDowns.DrawRequestNumber"
    Flag =0
    Expression ="DrawDowns.DrawNumber"
    Flag =0
    Expression ="Units.UnitInvoiceNum"
    Flag =0
    Expression ="Units.UnitId"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="DrawDowns.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="4185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Draw"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RequestNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescDocVer"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawRequestNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =34
    Top =20
    Right =1344
    Bottom =813
    Left =-1
    Top =-1
    Right =1278
    Bottom =544
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =746
        Top =45
        Right =929
        Bottom =420
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =484
        Top =45
        Right =685
        Bottom =407
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =233
        Top =38
        Right =435
        Bottom =537
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =484
        Top =422
        Right =681
        Bottom =566
        Top =0
        Name ="Vendors"
        Name =""
    End
End
