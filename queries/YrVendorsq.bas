Operation =1
Option =2
Where ="(((Units.UnitStatus)=\"b\")) OR (((Units.UnitStatus)=\"a\") AND ((Units.UnitAcce"
    "ptDate) Is Not Null And (Units.UnitAcceptDate)>#12/30/2010#)) OR (((Units.UnitSt"
    "atus)=\"a\") AND ((Units.UnitInServ) Is Not Null And (Units.UnitInServ)>#12/30/2"
    "010#))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.VendorREF"
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
        dbText "Name" ="Units.VendorREF"
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
    Bottom =253
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
