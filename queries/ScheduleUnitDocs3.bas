Operation =1
Option =1
Begin InputTables
    Name ="vw_ScheduleUnitDoc"
    Name ="Schedule_Group_WorkTable"
End
Begin OutputColumns
End
Begin Joins
    LeftTable ="vw_ScheduleUnitDoc"
    RightTable ="Schedule_Group_WorkTable"
    Expression ="vw_ScheduleUnitDoc.SchID = Schedule_Group_WorkTable.ScheduleID"
    Flag =1
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
        dbText "Name" ="vw_ScheduleUnitDoc.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.InService"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.EstiDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.MstrLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.InterimBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Unit No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.TitledState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.InterimDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Co_Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Lease No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_Group_WorkTable.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.ActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.StreetAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.County"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.Title Num"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.CustodianID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleUnitDoc.InterimAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_Group_WorkTable.ScheduleId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_Group_WorkTable.GroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =200
    Right =2007
    Bottom =1198
    Left =-1
    Top =-1
    Right =1895
    Bottom =135
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =327
        Top =1
        Right =471
        Bottom =145
        Top =0
        Name ="vw_ScheduleUnitDoc"
        Name =""
    End
    Begin
        Left =122
        Top =2
        Right =236
        Bottom =117
        Top =0
        Name ="Schedule_Group_WorkTable"
        Name =""
    End
End
