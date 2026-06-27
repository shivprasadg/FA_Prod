Operation =1
Option =0
Where ="(((UnitInspections.INSP) Is Null))"
Begin InputTables
    Name ="SELECT CLng(Nz([REQUEST ID],0)) AS RequestID, NewDekraInspections.STATUS, NewDek"
        "raInspections.[CLAIM#] AS ClaimNum, NewDekraInspections.[UNIT#] AS iUnitNum, CDa"
        "te(Nz([NewDekraInspections.REQUEST],[NewDekraInspections.REQUEST])) AS iDate, CL"
        "ng(Nz([INSP #],0)) AS INSP, NewDekraInspections.LinkAddr, CCur(Nz([PRICE TO CUST"
        "OMER],0)) AS iCost, NewDekraInspections.VIN, NewDekraInspections.TYPE, NewDekraI"
        "nspections.Year, NewDekraInspections.MAKE, NewDekraInspections.MODEL, NewDekraIn"
        "spections.CITY, NewDekraInspections.ST, NewDekraInspections.LOCATION FROM NewDek"
        "raInspections WHERE (((NewDekraInspections.STATUS)=\"Completed\"))"
    Alias ="NewInsp"
    Name ="Units"
    Name ="UnitInspections"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="NewInsp.RequestID"
    Expression ="NewInsp.STATUS"
    Expression ="NewInsp.ClaimNum"
    Expression ="NewInsp.iUnitNum"
    Expression ="NewInsp.iDate"
    Expression ="NewInsp.INSP"
    Expression ="NewInsp.LinkAddr"
    Expression ="NewInsp.iCost"
    Expression ="NewInsp.VIN"
    Expression ="NewInsp.TYPE"
    Expression ="NewInsp.Year"
    Expression ="NewInsp.MAKE"
    Expression ="NewInsp.MODEL"
    Expression ="NewInsp.CITY"
    Expression ="NewInsp.ST"
    Expression ="NewInsp.LOCATION"
End
Begin Joins
    LeftTable ="NewInsp"
    RightTable ="Units"
    Expression ="NewInsp.VIN = Units.UnitVIN"
    Flag =1
    LeftTable ="NewInsp"
    RightTable ="Units"
    Expression ="NewInsp.iUnitNum = Units.UnitUnitNum"
    Flag =1
    LeftTable ="NewInsp"
    RightTable ="UnitInspections"
    Expression ="NewInsp.INSP = UnitInspections.INSP"
    Flag =2
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
        dbText "Name" ="NewInsp.ClaimNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.RequestID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.iUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.iDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.INSP"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NewInsp.iCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.STATUS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.MODEL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.LOCATION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.LinkAddr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="7980"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NewInsp.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.MAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.CITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewInsp.ST"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =28
    Right =1426
    Bottom =1021
    Left =-1
    Top =-1
    Right =1324
    Bottom =667
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =357
        Top =0
        Name ="NewInsp"
        Name =""
    End
    Begin
        Left =245
        Top =48
        Right =389
        Bottom =557
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =583
        Top =70
        Right =727
        Bottom =214
        Top =0
        Name ="UnitInspections"
        Name =""
    End
End
