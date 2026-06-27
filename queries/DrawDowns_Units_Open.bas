Operation =1
Option =0
Where ="(((DrawDowns.DrawDownId) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="DrawDowns"
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
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.UnitNum"
    Alias ="UnitDesc"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Expression ="vw_SixKeys.UnitCost"
    Alias ="OEC"
    Expression ="vw_SixKeys.OECUnit"
    Alias ="Invoice"
    Expression ="vw_SixKeys.UnitInvoiceNum"
    Alias ="InvoiceDate"
    Expression ="vw_SixKeys.UnitinvoiceDate"
    Expression ="vw_SixKeys.UnitDescDocVer"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="vw_SixKeys.VendorId"
    Expression ="vw_SixKeys.MLOrig"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="vw_SixKeys"
    Expression ="DrawDowns.DrawDownId = vw_SixKeys.DrawDownId"
    Flag =3
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
    Expression ="vw_SixKeys.UnitInvoiceNum"
    Flag =0
    Expression ="vw_SixKeys.UnitId"
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
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceDate"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="vw_SixKeys.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.OECUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =120
    Top =93
    Right =1430
    Bottom =1091
    Left =-1
    Top =-1
    Right =1286
    Bottom =561
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =685
        Top =26
        Right =868
        Bottom =401
        Top =0
        Name ="DrawDowns"
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
        Left =472
        Top =325
        Right =669
        Bottom =469
        Top =0
        Name ="Vendors"
        Name =""
    End
End
