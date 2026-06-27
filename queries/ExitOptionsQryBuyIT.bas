Operation =1
Option =8
Where ="(((UnitExitOption.OptionTypeID)=0))"
Begin InputTables
    Name ="EXITOptions_WorkTable"
    Name ="UnitExitOption"
    Name ="Units"
    Name ="Units"
    Alias ="Units_1"
End
Begin OutputColumns
    Expression ="UnitExitOption.*"
    Expression ="Units.UnitUnitNum"
    Expression ="Units_1.UnitPDStartDate"
    Alias ="ParkDateProjected"
    Expression ="DLookUp(\"ParkDateProjected\",\"RMKT\",\"UnitRef=\" & [Units].[UnitID])"
    Alias ="AdjBookValue"
    Expression ="CCur(Nz(DLookUp(\"AdjBookValue\",\"RMKT\",\"UnitRef=\" & [Units].[UnitID]),0))"
    Alias ="TargetSale"
    Expression ="CCur(Nz(DLookUp(\"TargetSale\",\"RMKT\",\"UnitRef=\" & [Units].[UnitID]),0))"
    Alias ="BuyOutCost"
    Expression ="CCur(Nz(DLookUp(\"BuyOutCost\",\"RMKT\",\"UnitRef=\" & [Units].[UnitID]),0))"
    Alias ="TermPenalty"
    Expression ="CCur(Nz(DLookUp(\"TermPenalty\",\"RMKT\",\"UnitRef=\" & [Units].[UnitID]),0))"
End
Begin Joins
    LeftTable ="EXITOptions_WorkTable"
    RightTable ="UnitExitOption"
    Expression ="EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId"
    Flag =1
    LeftTable ="UnitExitOption"
    RightTable ="Units"
    Expression ="UnitExitOption.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="UnitExitOption"
    RightTable ="Units_1"
    Expression ="UnitExitOption.ReplacedByUnitID = Units_1.UnitID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AdjBookValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.LeaseOptionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.StartOrPurposedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrginalBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrigLeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrginalEOLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeReplied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkDateProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.FinalRentAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.FinalRentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeNotified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Fees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UserName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UIDX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrigResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateEarly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.DatePosted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EarlyGapOverLapMsg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_1.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.CostGainOfSale"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-123
    Top =26
    Right =1346
    Bottom =1007
    Left =-1
    Top =-1
    Right =1437
    Bottom =629
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="EXITOptions_WorkTable"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =420
        Top =217
        Right =564
        Bottom =361
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =242
        Top =215
        Right =386
        Bottom =359
        Top =0
        Name ="Units_1"
        Name =""
    End
End
