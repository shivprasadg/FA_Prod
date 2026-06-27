Operation =1
Option =0
Where ="(((FleetStudyImport.[VIN]) In (SELECT [VIN] FROM [FleetStudyImport] As Tmp GROUP"
    " BY [VIN] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="FleetStudyImport"
End
Begin OutputColumns
    Expression ="FleetStudyImport.[VIN]"
    Expression ="FleetStudyImport.[UnitNumber]"
    Expression ="FleetStudyImport.[UnitYear]"
    Expression ="FleetStudyImport.[UnitMake]"
    Expression ="FleetStudyImport.[UnitModel]"
End
Begin OrderBy
    Expression ="FleetStudyImport.[VIN]"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1504
    Bottom =1038
    Left =-1
    Top =-1
    Right =1472
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="FleetStudyImport"
        Name =""
    End
End
