Operation =1
Option =8
Begin InputTables
    Name ="UnitExitOption"
    Name ="Units"
    Name ="EXITOptions_WorkTable"
End
Begin OutputColumns
    Expression ="Units.GroupREF"
    Expression ="Units.UnitUnitNum"
    Expression ="UnitExitOption.LeaseOptionID"
    Expression ="UnitExitOption.UnitID"
    Expression ="UnitExitOption.OptionTypeID"
    Expression ="UnitExitOption.OptionTypeDesc"
    Expression ="UnitExitOption.OptionSubType"
    Alias ="Start-Purposed"
    Expression ="UnitExitOption.StartOrPurposedDate"
    Alias ="End-Actual"
    Expression ="UnitExitOption.EndOrActualDate"
    Expression ="UnitExitOption.TermMonths"
    Expression ="UnitExitOption.Rent"
    Expression ="UnitExitOption.Fees"
    Expression ="UnitExitOption.Residual"
    Expression ="UnitExitOption.ReplacedByScheduleID"
    Expression ="UnitExitOption.ReplacedByGroupID"
    Expression ="UnitExitOption.ReplacedByUnitID"
    Expression ="UnitExitOption.TerminateDate"
    Expression ="UnitExitOption.TerminateFee"
    Expression ="UnitExitOption.FinalRentAmt"
    Expression ="UnitExitOption.FinalRentDate"
    Expression ="UnitExitOption.AssigneeNotified"
    Expression ="UnitExitOption.AssigneeReplied"
    Expression ="UnitExitOption.OrginalBLCD"
    Expression ="UnitExitOption.OrigLeaseTerm"
    Expression ="UnitExitOption.OrginalEOLDate"
    Expression ="UnitExitOption.OrigResidual"
    Expression ="UnitExitOption.TerminateEarly"
    Expression ="UnitExitOption.UserName"
    Expression ="UnitExitOption.DatePosted"
    Expression ="UnitExitOption.Notes"
    Expression ="UnitExitOption.UIDX"
    Alias ="NewSchedule"
    Expression ="IIf(Nz([ReplacedByScheduleID],0)=0,Null,DLookUp(\"SchNo\",\"Schedule\",\"SchId=\""
        " & [ReplacedByScheduleID] & \"\"))"
    Alias ="NewGroup"
    Expression ="IIf(Nz([ReplacedByGroupID],0)=0,Null,DLookUp(\"sGrpNo\",\"SchGrp\",\"SgrpID=\" &"
        " [ReplacedByGroupID] & \"\"))"
    Alias ="ExchangeITDate"
    Expression ="IIf([OptionTypeID]=1,Nz([End-Actual],[Start-Purposed]),Null)"
    Alias ="ExchangeITStatus"
    Expression ="IIf([OptionTypeID]<>1,Null,IIf(IsNull([End-Actual]),\"P\",\"A\"))"
    Alias ="ExchangeITSchGrpUnit"
    Expression ="IIf(IsNull([NewSchedule]),Null,[NewSchedule] & \"|\" & [NewGroup] & IIf(IsNull(["
        "ReplacedByUnitID]),\"|----\",\"|\" & [ReplacedByUnitID]))"
    Alias ="ExtendITStartDate"
    Expression ="IIf([OptionTypeID]>2,[StartOrPurposedDate],Null)"
    Alias ="ExtendITEndDate"
    Expression ="IIf([OptionTypeID]>2,[EndOrActualDate],Null)"
    Alias ="ExtendedFees"
    Expression ="IIf([OptionTypeID]>2,[Fees],Null)"
    Alias ="ExtendedRent"
    Expression ="IIf([OptionTypeID]>2,[Rent],Null)"
    Alias ="TermDate"
    Expression ="IIf([OptionTypeID]=2,Nz([End-Actual],[Start-Purposed]),Null)"
    Alias ="TermStatus"
    Expression ="IIf([OptionTypeID]<>2,Null,IIf(IsNull([End-Actual]),\"P\",\"A\"))"
    Alias ="GainLoss"
    Expression ="IIf([OptionTypeID]<>2,Null,Format([TerminateFee],\"$#,###.##\"))"
    Alias ="TermEarly"
    Expression ="IIf([OptionTypeID]<>2,Null,[TerminateEarly])"
End
Begin Joins
    LeftTable ="UnitExitOption"
    RightTable ="Units"
    Expression ="UnitExitOption.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="UnitExitOption"
    RightTable ="EXITOptions_WorkTable"
    Expression ="UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="UnitExitOption.UnitID"
    Flag =0
    Expression ="UnitExitOption.OptionTypeDesc"
    Flag =0
    Expression ="UnitExitOption.StartOrPurposedDate"
    Flag =1
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
        dbText "Name" ="UnitExitOption.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.LeaseOptionID"
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
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeNotified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByScheduleID"
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
        dbText "Name" ="UnitExitOption.DatePosted"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitExitOption.Notes"
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
        dbText "Name" ="UnitExitOption.OrigResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateEarly"
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
        dbText "Name" ="UnitExitOption.OptionSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeReplied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="End-Actual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Start-Purposed"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITSchGrpUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendITStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendITEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendedFees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermEarly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITStatus"
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
        dbText "Name" ="GainLoss"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =185
    Top =115
    Right =1802
    Bottom =1066
    Left =-1
    Top =-1
    Right =1593
    Bottom =163
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =251
        Top =29
        Right =430
        Bottom =256
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =466
        Top =94
        Right =610
        Bottom =238
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =36
        Top =45
        Right =180
        Bottom =189
        Top =0
        Name ="EXITOptions_WorkTable"
        Name =""
    End
End
