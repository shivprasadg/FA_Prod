Operation =1
Option =0
Begin InputTables
    Name ="FALeased_Unitcount"
    Name ="ByYear_UnitsDelivered"
    Name ="ByYear_UnitsOrdered-PO"
    Name ="ByYear_UnitsTitled"
    Name ="ByYear_UnitsAccepted"
    Name ="ByYear_UnitsOffLease"
    Name ="ByYear_UnitsInServiced"
End
Begin OutputColumns
    Expression ="FALeased_Unitcount.ClientShortName"
    Alias ="Ordered"
    Expression ="CDbl(Nz([UnitOrdered],0))"
    Alias ="Titled"
    Expression ="CDbl(Nz([UnitTitled],0))"
    Alias ="Delivered"
    Expression ="CDbl(Nz([UnitDelivered],0))"
    Alias ="Accepted"
    Expression ="CDbl(Nz([UnitAccepted],0))"
    Alias ="InServiced"
    Expression ="CDbl(Nz([UnitsInServiced],0))"
    Alias ="OffLease"
    Expression ="CDbl(Nz([UnitsOffLease],0))"
    Expression ="[ByYear_UnitsOrdered-PO].FirstPO"
    Expression ="[ByYear_UnitsOrdered-PO].LastPO"
    Expression ="ByYear_UnitsInServiced.LastInServiced"
    Alias ="LastOffLease"
    Expression ="ByYear_UnitsOffLease.LastOL"
    Expression ="FALeased_Unitcount.ClientGroupId"
    Expression ="FALeased_Unitcount.clientID"
End
Begin Joins
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsOrdered-PO"
    Expression ="FALeased_Unitcount.clientID = [ByYear_UnitsOrdered-PO].clientID"
    Flag =2
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsDelivered"
    Expression ="FALeased_Unitcount.clientID = ByYear_UnitsDelivered.clientID"
    Flag =2
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsTitled"
    Expression ="FALeased_Unitcount.clientID = ByYear_UnitsTitled.clientID"
    Flag =2
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsAccepted"
    Expression ="FALeased_Unitcount.clientID = ByYear_UnitsAccepted.clientID"
    Flag =2
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsOffLease"
    Expression ="FALeased_Unitcount.clientID = ByYear_UnitsOffLease.clientID"
    Flag =2
    LeftTable ="FALeased_Unitcount"
    RightTable ="ByYear_UnitsInServiced"
    Expression ="FALeased_Unitcount.clientID = ByYear_UnitsInServiced.clientID"
    Flag =2
End
Begin OrderBy
    Expression ="FALeased_Unitcount.ClientShortName"
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
        dbText "Name" ="FALeased_Unitcount.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ordered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiced"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Titled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delivered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Accepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ByYear_UnitsOrdered-PO].FirstPO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ByYear_UnitsOrdered-PO].LastPO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ByYear_UnitsInServiced.LastInServiced"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastOffLease"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =35
    Top =49
    Right =1630
    Bottom =1047
    Left =-1
    Top =-1
    Right =1563
    Bottom =646
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =69
        Top =-1
        Right =213
        Bottom =212
        Top =0
        Name ="FALeased_Unitcount"
        Name =""
    End
    Begin
        Left =598
        Top =172
        Right =742
        Bottom =316
        Top =0
        Name ="ByYear_UnitsDelivered"
        Name =""
    End
    Begin
        Left =282
        Top =18
        Right =426
        Bottom =162
        Top =0
        Name ="ByYear_UnitsOrdered-PO"
        Name =""
    End
    Begin
        Left =442
        Top =74
        Right =586
        Bottom =218
        Top =0
        Name ="ByYear_UnitsTitled"
        Name =""
    End
    Begin
        Left =745
        Top =267
        Right =889
        Bottom =482
        Top =0
        Name ="ByYear_UnitsAccepted"
        Name =""
    End
    Begin
        Left =969
        Top =526
        Right =1113
        Bottom =670
        Top =0
        Name ="ByYear_UnitsOffLease"
        Name =""
    End
    Begin
        Left =901
        Top =383
        Right =1045
        Bottom =527
        Top =0
        Name ="ByYear_UnitsInServiced"
        Name =""
    End
End
