Operation =1
Option =0
Where ="(((FARentBillAssignedQ.GroupID)=[Forms]![Billing]![GroupId]))"
Begin InputTables
    Name ="FARentBillAssignedQ"
End
Begin OutputColumns
    Expression ="FARentBillAssignedQ.*"
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
        dbText "Name" ="FARentBillq.sgrpid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocationCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BLocZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.UnitsBillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.unittax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Schedule.invfld1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bill1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Expr1049"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Schedule.invfld1desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bill2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.TaxAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Schedule.invfld2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Schedule.invfld2desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.RentNoTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.FSLTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.UnitTotalRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BillToAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.DueDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.BankContract"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1228
    Bottom =908
    Left =-1
    Top =-1
    Right =1196
    Bottom =570
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =74
        Top =95
        Right =372
        Bottom =413
        Top =0
        Name ="FARentBillAssignedQ"
        Name =""
    End
End
