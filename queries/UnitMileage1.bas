Operation =1
Option =0
Begin InputTables
    Name ="vw_UMP_Mileage"
    Name ="MileageOdoReadingAdj"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="vw_UMP_Mileage.UnitID"
    Alias ="UnitNo"
    Expression ="vw_UMP_Mileage.UnitNumber"
    Alias ="Odometer"
    Expression ="Max(IIf([OBCDataStatus]=6,0,[OdometerEnd]))"
    Alias ="MPGOverAll"
    Expression ="Round(IIf(Nz(Max([MilesLTD]),0)>0 And Nz(Max([FuelLTD]),0)>0,(Max([MilesLTD])/Ma"
        "x([FuelLTD])),0),2)"
    Alias ="MPGMoving"
    Expression ="Round(IIf(Nz(Max([MilesLTD]),0)>0 And Nz(Max([FuelmovingLTD]),0)>0,(Max([MilesLT"
        "D])/Max([FuelMovingLTD])),0),2)"
    Alias ="ODOReadDate"
    Expression ="Max(vw_UMP_Mileage.LogActivityDate)"
    Alias ="MonthsReported"
    Expression ="Count(vw_UMP_Mileage.UnitId)"
    Alias ="OBCStatus"
    Expression ="Max(vw_UMP_Mileage.OBCDataStatus)"
    Alias ="OdoError"
    Expression ="0"
    Alias ="MPGError"
    Expression ="Max(IIf([OBCDataStatus]=7,[vw_UMP_Mileage].[MPGMoving],0))"
    Expression ="MileageOdoReadingAdj.FirstReport"
    Expression ="MileageOdoReadingAdj.LastReport"
    Expression ="MileageOdoReadingAdj.ODOMiles"
    Expression ="MileageOdoReadingAdj.MthsRptd"
    Expression ="MileageOdoReadingAdj.MPM"
End
Begin Joins
    LeftTable ="vw_UMP_Mileage"
    RightTable ="MileageOdoReadingAdj"
    Expression ="vw_UMP_Mileage.UnitID = MileageOdoReadingAdj.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_UMP_Mileage.UnitID"
    Flag =1
End
Begin Groups
    Expression ="vw_UMP_Mileage.UnitID"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.UnitNumber"
    GroupLevel =0
    Expression ="MileageOdoReadingAdj.FirstReport"
    GroupLevel =0
    Expression ="MileageOdoReadingAdj.LastReport"
    GroupLevel =0
    Expression ="MileageOdoReadingAdj.ODOMiles"
    GroupLevel =0
    Expression ="MileageOdoReadingAdj.MthsRptd"
    GroupLevel =0
    Expression ="MileageOdoReadingAdj.MPM"
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
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODOReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthsReported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OdoError"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MPGError"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MPGMoving"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MPGOverAll"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileageOdoReadingAdj.ODOMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileageOdoReadingAdj.LastReport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileageOdoReadingAdj.FirstReport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileageOdoReadingAdj.MPM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileageOdoReadingAdj.MthsRptd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =1
    Top =71
    Right =1249
    Bottom =886
    Left =-1
    Top =-1
    Right =1216
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =363
        Top =22
        Right =672
        Bottom =337
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
    Begin
        Left =743
        Top =78
        Right =887
        Bottom =331
        Top =0
        Name ="MileageOdoReadingAdj"
        Name =""
    End
End
