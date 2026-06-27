Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposalUnits"
    Name ="RemarketingReportAllUnits"
    Name ="LookUp_UnitSubTypes"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.UnitId"
    Alias ="Desc"
    Expression ="Nz([UnitSubTypeAbbr],'') & ' - ' & [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitMode"
        "l] & \" \" & [Axle]"
    Alias ="STS"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="BuyersProposalUnits.ProposalID"
    Alias ="Insp"
    Expression ="IIf(IsNull([InspectionLinkWeb]),\"No\",\"Yes\")"
    Expression ="RemarketingReportAllUnits.InspectionLinkWeb"
    Alias ="Target"
    Expression ="Format([TargetSalePrice],\"$#,###\")"
    Alias ="ODO"
    Expression ="RemarketingReportAllUnits.ODORead"
    Expression ="RemarketingReportAllUnits.ODOSource"
    Alias ="DaysIn"
    Expression ="RemarketingReportAllUnits.DaysInInventory"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.Schedule"
    Alias ="Location"
    Expression ="[LCity] & \", \" & [LState] & \". \" & [LZip]"
    Expression ="RemarketingReportAllUnits.LocationID"
    Expression ="RemarketingReportAllUnits.Latitude"
    Expression ="RemarketingReportAllUnits.Longitude"
End
Begin Joins
    LeftTable ="BuyersProposalUnits"
    RightTable ="RemarketingReportAllUnits"
    Expression ="BuyersProposalUnits.UnitID = RemarketingReportAllUnits.UnitID"
    Flag =1
    LeftTable ="RemarketingReportAllUnits"
    RightTable ="LookUp_UnitSubTypes"
    Expression ="RemarketingReportAllUnits.UnitSubType = LookUp_UnitSubTypes.UnitSubType"
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
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="STS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Insp"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InspectionLinkWeb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ODOSource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalUnits.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Target"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SourceCompany"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="4410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Longitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Latitude"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =182
    Top =59
    Right =1708
    Bottom =1057
    Left =-1
    Top =-1
    Right =1494
    Bottom =544
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="BuyersProposalUnits"
        Name =""
    End
    Begin
        Left =325
        Top =48
        Right =526
        Bottom =553
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
    Begin
        Left =667
        Top =401
        Right =811
        Bottom =545
        Top =0
        Name ="LookUp_UnitSubTypes"
        Name =""
    End
End
