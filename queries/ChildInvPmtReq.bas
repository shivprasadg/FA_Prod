Operation =1
Option =0
Where ="(((MstrLease.MLNo)=[enter Master Lease No]) AND ((ChildParts.ChildInvoiceNum)=[e"
    "nter invoice no]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildParts"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitVIN"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorAddress"
    Expression ="Vendors.VendorCity"
    Expression ="Vendors.VendorState"
    Expression ="Vendors.VendorZip"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="Vendors.VendorPmtType"
    Expression ="ChildParts.childFinalCost"
    Expression ="Vendors.VENDORPAYEE"
    Expression ="Vendors.VENDORABA"
    Expression ="Vendors.VENDORACH"
    Expression ="Vendors.VENDORBKACCT"
    Expression ="Vendors.VENDORBANK"
    Expression ="Vendors.VENDORBADD1"
    Expression ="Vendors.VEDNORBADD2"
    Expression ="Vendors.VENDORBCITY"
    Expression ="Vendors.VENDORBST"
    Expression ="Vendors.VENDORBZIP"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.ChildVIN"
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
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="Vendors"
    Expression ="ChildParts.ChildVendorREF = Vendors.VendorID"
    Flag =2
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
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
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPmtType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORPAYEE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBKACCT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBANK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBADD1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VEDNORBADD2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBZIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
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
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
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
    Bottom =158
    Left =192
    Top =0
    ColumnsShown =539
    Begin
        Left =-161
        Top =-6
        Right =-17
        Bottom =138
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =9
        Top =3
        Right =153
        Bottom =147
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =212
        Top =11
        Right =356
        Bottom =155
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =422
        Top =10
        Right =566
        Bottom =154
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =463
        Top =4
        Right =607
        Bottom =148
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =654
        Top =2
        Right =797
        Bottom =146
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =837
        Top =18
        Right =981
        Bottom =162
        Top =0
        Name ="Vendors"
        Name =""
    End
End
