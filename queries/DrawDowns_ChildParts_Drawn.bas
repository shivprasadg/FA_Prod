Operation =1
Option =0
Where ="(((ChildParts.DrawDownId) Is Not Null) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Vendors"
    Name ="vw_SixKeys"
    Name ="ChildParts"
    Name ="DrawDowns"
    Name ="Banks"
End
Begin OutputColumns
    Expression ="ChildParts.ChildPartID"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.UnitNum"
    Alias ="CInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="ChildParts.ChildType"
    Alias ="CCost"
    Expression ="ChildParts.childFinalCost"
    Expression ="ChildParts.DrawDownId"
    Alias ="cIFBank"
    Expression ="Banks.BankShortName"
    Expression ="ChildParts.cIFAmount"
    Alias ="CVendor"
    Expression ="Vendors.vendorsnm"
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Alias ="RequestNumber"
    Expression ="DrawDowns.DrawRequestNumber"
    Alias ="ChildDesc"
    Expression ="[ChildMake] & \" \" & [ChildClass] & \" \" & [ChildType]"
    Expression ="DrawDowns.DrawName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="Vendors.VendorID"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="DrawDowns.DrawRequestNumber"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="ChildParts"
    Expression ="vw_SixKeys.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="Vendors"
    RightTable ="ChildParts"
    Expression ="Vendors.VendorID = ChildParts.ChildVendorREF"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="DrawDowns"
    Expression ="vw_SixKeys.DrawDownId = DrawDowns.DrawDownId"
    Flag =2
    LeftTable ="ChildParts"
    RightTable ="Banks"
    Expression ="ChildParts.cIFBankID = Banks.BankID"
    Flag =2
End
Begin OrderBy
    Expression ="Vendors.vendorsnm"
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
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CVendor"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RequestNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildDesc"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.assetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawRequestNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.cIFAmount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =112
    Right =1911
    Bottom =951
    Left =-1
    Top =-1
    Right =1870
    Bottom =459
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =71
        Top =34
        Right =215
        Bottom =178
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =472
        Top =44
        Right =662
        Bottom =321
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =250
        Top =45
        Right =430
        Bottom =410
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =705
        Top =37
        Right =849
        Bottom =382
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =64
        Top =216
        Right =208
        Bottom =425
        Top =0
        Name ="Banks"
        Name =""
    End
End
