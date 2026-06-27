Operation =1
Option =2
Having ="(((ClientsFAUnits.ClientGroupName) Is Not Null))"
Begin InputTables
    Name ="ClientsFAUnits"
    Name ="ClientsNonFAUnits"
    Name ="ClientsFAUnitsFirstOrder"
End
Begin OutputColumns
    Expression ="ClientsFAUnits.ClientGroupName"
    Expression ="ClientsFAUnits.ClientShortName"
    Alias ="FirstDeliveryDate"
    Expression ="ClientsFAUnitsFirstOrder.ActualDeliveryDate"
    Alias ="SumOfFAUnits"
    Expression ="Sum(ClientsFAUnits.FAUnits)"
    Alias ="NotFAUnits"
    Expression ="Max((CInt(Nz([NonFAUnits],0))))"
    Alias ="TotalUnits"
    Expression ="Sum([FAUnits])+Nz(Max([NonFAUnits]))"
End
Begin Joins
    LeftTable ="ClientsFAUnitsFirstOrder"
    RightTable ="ClientsFAUnits"
    Expression ="ClientsFAUnitsFirstOrder.clientID = ClientsFAUnits.clientID"
    Flag =2
    LeftTable ="ClientsFAUnitsFirstOrder"
    RightTable ="ClientsNonFAUnits"
    Expression ="ClientsFAUnitsFirstOrder.clientID = ClientsNonFAUnits.clientID"
    Flag =2
End
Begin Groups
    Expression ="ClientsFAUnits.ClientGroupName"
    GroupLevel =0
    Expression ="ClientsFAUnits.ClientShortName"
    GroupLevel =0
    Expression ="ClientsFAUnitsFirstOrder.ActualDeliveryDate"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="ClientsFAUnits.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientsFAUnits.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NotFAUnits"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="SumOfFAUnits"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="FirstDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientsFAUnitsFirstOrder.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-332
    Top =126
    Right =1068
    Bottom =1059
    Left =-1
    Top =-1
    Right =1382
    Bottom =571
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =315
        Top =23
        Right =459
        Bottom =304
        Top =0
        Name ="ClientsFAUnits"
        Name =""
    End
    Begin
        Left =496
        Top =119
        Right =640
        Bottom =394
        Top =0
        Name ="ClientsNonFAUnits"
        Name =""
    End
    Begin
        Left =61
        Top =32
        Right =205
        Bottom =322
        Top =0
        Name ="ClientsFAUnitsFirstOrder"
        Name =""
    End
End
