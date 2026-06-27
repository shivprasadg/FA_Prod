Operation =1
Option =0
Where ="(((UnitVinDecoder.UnitID) In (SELECT [UnitID] FROM [UnitVinDecoder] As Tmp GROUP"
    " BY [UnitID] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="UnitVinDecoder"
End
Begin OutputColumns
    Expression ="UnitVinDecoder.UnitID"
    Expression ="UnitVinDecoder.DateDecoded"
    Expression ="UnitVinDecoder.status"
    Expression ="UnitVinDecoder.ClientID"
    Expression ="UnitVinDecoder.VIN"
    Expression ="UnitVinDecoder.make"
    Expression ="UnitVinDecoder.modelyear"
    Expression ="UnitVinDecoder.INDX"
End
Begin OrderBy
    Expression ="UnitVinDecoder.UnitID"
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
        dbText "Name" ="UnitVinDecoder.status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.[DateDecoded]"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.[status]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.[ClientID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.INDX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.DateDecoded"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.modelyear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =266
    Top =8
    Right =1456
    Bottom =989
    Left =-1
    Top =-1
    Right =1172
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UnitVinDecoder"
        Name =""
    End
End
