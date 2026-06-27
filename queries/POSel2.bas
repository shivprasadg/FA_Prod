Operation =1
Option =2
Where ="(((Vendors.VendorName) Is Not Null) AND ((MstrLease.MLOrig)=\"fa\") AND ((SchGrp"
    ".AssignmentDate) Is Null))"
Begin InputTables
    Name ="ChildParts"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchID"
    Expression ="Vendors.VendorName"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpSchID = Units.GroupREF"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="Units"
    Expression ="ChildParts.UnitREF = Units.UnitID"
    Flag =3
    LeftTable ="ChildParts"
    RightTable ="Vendors"
    Expression ="ChildParts.ChildVendorREF = Vendors.VendorID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Vendors.VendorName"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =392
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =375
        Top =325
        Right =519
        Bottom =426
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =405
        Top =-1
        Right =549
        Bottom =143
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =570
        Top =-5
        Right =714
        Bottom =139
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =371
        Top =172
        Right =515
        Bottom =316
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =576
        Top =298
        Right =720
        Bottom =440
        Top =0
        Name ="Vendors"
        Name =""
    End
End
