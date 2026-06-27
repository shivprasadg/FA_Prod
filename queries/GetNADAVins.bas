Operation =1
Option =0
Where ="(((Len([six.VIN]))=17) AND ((UnitVinDetails.status)=\"0\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Alias ="Six"
    Name ="SELECT ClientID, VinGroup FROM vw_SixKeys GROUP BY ClientID, VinGroup HAVING (vw"
        "_SixKeys.[ClientID])=4229"
    Alias ="V"
    Name ="UnitVinDetails"
End
Begin OutputColumns
    Alias ="VIN"
    Expression ="Last(Six.Vin)"
    Expression ="Six.Type"
    Expression ="V.VinGroup"
    Alias ="ClientID"
    Expression ="[vw_SixKeys].[ClientID]"
    Alias ="ModelYear"
    Expression ="Max(Six.Myear)"
End
Begin Joins
    LeftTable ="Six"
    RightTable ="V"
    Expression ="Six.VinGRoup = V.VinGroup"
    Flag =3
    LeftTable ="Six"
    RightTable ="UnitVinDetails"
    Expression ="Six.VIN = UnitVinDetails.VIN"
    Flag =2
End
Begin Groups
    Expression ="Six.Type"
    GroupLevel =0
    Expression ="V.VinGroup"
    GroupLevel =0
    Expression ="[vw_SixKeys].[ClientID]"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="V.VinGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Six.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =136
    Top =4
    Right =1351
    Bottom =982
    Left =-1
    Top =-1
    Right =1197
    Bottom =633
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =514
        Top =72
        Right =658
        Bottom =464
        Top =0
        Name ="Six"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =248
        Top =0
        Name ="V"
        Name =""
    End
    Begin
        Left =699
        Top =328
        Right =843
        Bottom =472
        Top =0
        Name ="UnitVinDetails"
        Name =""
    End
End
