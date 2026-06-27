Operation =1
Option =0
Begin InputTables
    Name ="Client_LPAF"
End
Begin OutputColumns
    Expression ="Client_LPAF.Schedule"
    Expression ="Client_LPAF.UnitGroup"
    Alias ="CountOfUnitID"
    Expression ="Count(Client_LPAF.UnitID)"
    Expression ="Client_LPAF.Make"
    Expression ="Client_LPAF.Model"
    Expression ="Client_LPAF.Type"
    Alias ="SumOfUnit Cost"
    Expression ="Sum(Client_LPAF.[Unit Cost])"
    Alias ="SumOfChild Cost"
    Expression ="Sum(Client_LPAF.[Child Cost])"
    Expression ="Client_LPAF.[Domicile Location]"
    Expression ="Client_LPAF.[Target Delivery]"
    Expression ="Client_LPAF.[Spec Letter]"
    Expression ="Client_LPAF.[Annual Mileage]"
    Expression ="Client_LPAF.[LO Rent]"
    Expression ="Client_LPAF.[Primary Lease Term]"
    Expression ="Client_LPAF.[Synd Rent]"
    Expression ="Client_LPAF.[Lease Extend Term]"
    Expression ="Client_LPAF.[Lease Extend Rent]"
    Expression ="Client_LPAF.MTMRent"
    Expression ="Client_LPAF.SchID"
    Expression ="Client_LPAF.ClientShortName"
    Expression ="Client_LPAF.clientID"
End
Begin Groups
    Expression ="Client_LPAF.Schedule"
    GroupLevel =0
    Expression ="Client_LPAF.UnitGroup"
    GroupLevel =0
    Expression ="Client_LPAF.Make"
    GroupLevel =0
    Expression ="Client_LPAF.Model"
    GroupLevel =0
    Expression ="Client_LPAF.Type"
    GroupLevel =0
    Expression ="Client_LPAF.[Domicile Location]"
    GroupLevel =0
    Expression ="Client_LPAF.[Target Delivery]"
    GroupLevel =0
    Expression ="Client_LPAF.[Spec Letter]"
    GroupLevel =0
    Expression ="Client_LPAF.[Annual Mileage]"
    GroupLevel =0
    Expression ="Client_LPAF.[LO Rent]"
    GroupLevel =0
    Expression ="Client_LPAF.[Primary Lease Term]"
    GroupLevel =0
    Expression ="Client_LPAF.[Synd Rent]"
    GroupLevel =0
    Expression ="Client_LPAF.[Lease Extend Term]"
    GroupLevel =0
    Expression ="Client_LPAF.[Lease Extend Rent]"
    GroupLevel =0
    Expression ="Client_LPAF.MTMRent"
    GroupLevel =0
    Expression ="Client_LPAF.SchID"
    GroupLevel =0
    Expression ="Client_LPAF.ClientShortName"
    GroupLevel =0
    Expression ="Client_LPAF.clientID"
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
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SumOfChild Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_LPAF.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_LPAF.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnit Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_LPAF.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =24
    Right =1519
    Bottom =993
    Left =-1
    Top =-1
    Right =1459
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =81
        Top =40
        Right =225
        Bottom =513
        Top =0
        Name ="Client_LPAF"
        Name =""
    End
End
