Operation =1
Option =2
Where ="(((MnRData.LogActivityDate)=#3/1/2020#))"
Begin InputTables
    Name ="MnRData"
End
Begin OutputColumns
    Expression ="MnRData.UnitID"
    Expression ="MnRData.VIN"
    Alias ="PrevLogActivityDate"
    Expression ="MnRData.LogActivityDate"
    Alias ="PrevMilesLTD"
    Expression ="MnRData.MilesLTD"
    Alias ="PrevHoursLTD"
    Expression ="MnRData.HoursLTD"
    Alias ="PrevLiftGateLTD"
    Expression ="MnRData.LiftGateCyclesLTD"
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
    Right =1268
    Bottom =1271
    Left =-1
    Top =-1
    Right =1244
    Bottom =951
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MnRData"
        Name =""
    End
End
