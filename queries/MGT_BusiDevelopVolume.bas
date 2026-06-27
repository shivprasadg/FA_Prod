Operation =1
Option =0
Begin InputTables
    Name ="Awarded_Deals_Client"
    Name ="ML_Sch_Grp_FAID_Count_Cost_Rents"
End
Begin OutputColumns
    Expression ="Awarded_Deals_Client.FAID"
    Expression ="Awarded_Deals_Client.Client"
    Alias ="Units"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCount"
    Alias ="OEC"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalUnitCost"
    Expression ="Awarded_Deals_Client.LPAFQtr"
    Expression ="Awarded_Deals_Client.LPAFYear"
    Expression ="Awarded_Deals_Client.DealQtr"
    Expression ="Awarded_Deals_Client.DealYear"
    Expression ="Awarded_Deals_Client.DealUnitCount"
    Expression ="Awarded_Deals_Client.DealGrandTotal"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.OPSQtr"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.OPSYear"
    Alias ="ActualUnitCount"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCount"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalLORent"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalSyndRent"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCost"
    Expression ="ML_Sch_Grp_FAID_Count_Cost_Rents.ChildPartCost"
End
Begin Joins
    LeftTable ="Awarded_Deals_Client"
    RightTable ="ML_Sch_Grp_FAID_Count_Cost_Rents"
    Expression ="Awarded_Deals_Client.FAID = ML_Sch_Grp_FAID_Count_Cost_Rents.LFAID"
    Flag =2
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
dbMemo "Filter" ="([MGT_BusiDevelopVolume].[FAID]=\"13-507\")"
Begin
    Begin
        dbText "Name" ="Awarded_Deals_Client.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.DealGrandTotal"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.LPAFYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.DealYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.DealUnitCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.LPAFQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Awarded_Deals_Client.DealQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalLORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalSyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.TotalUnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualUnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.OPSQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ML_Sch_Grp_FAID_Count_Cost_Rents.OPSYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealUnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =224
    Top =48
    Right =1886
    Bottom =978
    Left =-1
    Top =-1
    Right =1630
    Bottom =579
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =113
        Top =95
        Right =257
        Bottom =472
        Top =0
        Name ="Awarded_Deals_Client"
        Name =""
    End
    Begin
        Left =369
        Top =133
        Right =513
        Bottom =499
        Top =0
        Name ="ML_Sch_Grp_FAID_Count_Cost_Rents"
        Name =""
    End
End
