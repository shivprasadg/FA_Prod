Operation =1
Option =0
Begin InputTables
    Name ="RemarketingUnitsReport"
    Name ="TrackedUnitLastMileage"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="TrackedUnitLastMileage.UnitID"
    Expression ="TrackedUnitLastMileage.MHDate"
    Expression ="TrackedUnitLastMileage.Odometer"
    Expression ="Clients.clientobc"
End
Begin Joins
    LeftTable ="RemarketingUnitsReport"
    RightTable ="TrackedUnitLastMileage"
    Expression ="RemarketingUnitsReport.AssetID = TrackedUnitLastMileage.UnitID"
    Flag =1
    LeftTable ="RemarketingUnitsReport"
    RightTable ="Clients"
    Expression ="RemarketingUnitsReport.clientID = Clients.clientID"
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
        dbText "Name" ="TrackedUnitLastMileage.MHDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrackedUnitLastMileage.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientobc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =173
    Top =46
    Right =1549
    Bottom =861
    Left =-1
    Top =-1
    Right =1344
    Bottom =464
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =17
        Top =3
        Right =174
        Bottom =396
        Top =0
        Name ="RemarketingUnitsReport"
        Name =""
    End
    Begin
        Left =213
        Top =267
        Right =357
        Bottom =411
        Top =0
        Name ="TrackedUnitLastMileage"
        Name =""
    End
    Begin
        Left =224
        Top =93
        Right =368
        Bottom =237
        Top =0
        Name ="Clients"
        Name =""
    End
End
