Operation =1
Option =0
Where ="(((FABankBill.LocationID)=1173) And ((FABankBill.SGrpID)=Forms!Billing!GroupId))"
Begin InputTables
    Name ="FABankBill"
End
Begin OutputColumns
    Expression ="FABankBill.*"
    Alias ="RentNoTax"
    Expression ="FABankBill.UNITRENT"
    Alias ="FSLTax"
    Expression ="IIf([UnitTaxExempt]=False,CCur([FSLFee]*[LTaxRate]),0)"
    Alias ="UnitTotalRent"
    Expression ="[RentNoTax]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0))"
    Alias ="BillToMBMAdr"
    Expression ="\"Meadowbrook Meat Company Inc.\" & Chr(13) & Chr(10) & \"PO BOX 6258\" & Chr(13"
        ") & Chr(10) & \"Rocky Mount NC 27802\" & Chr(13) & Chr(10) & \"Attn: Tiffany Hud"
        "son\""
    Alias ="InvoiceNumber"
    Expression ="FABankBill.BillNo"
    Alias ="BillToAddress"
    Expression ="FABankBill.BillToGroup"
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
        dbText "Name" ="FABankBill.Bills.BillID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocationID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
        dbInteger "ColumnWidth" ="3390"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocationName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="FABankBill.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="FABankBill.Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocationCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocationCounty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.stateRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocationZip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="FABankBill.BLoST3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="FABankBill.BLocStreet"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="FABankBill.BLocCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="FABankBill.BLocSt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="FABankBill.BLocZip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="FABankBill.lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.BillNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.Billdt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.Billfrom"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.Billto"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="FABankBill.vin4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="FABankBill.FSLFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UnitTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="FABankBill.UnitsBillRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="FABankBill.TaxAmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="FABankBill.Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="51"
    End
    Begin
        dbText "Name" ="FABankBill.Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="52"
    End
    Begin
        dbText "Name" ="FABankBill.BLocSt2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="55"
    End
    Begin
        dbText "Name" ="FABankBill.Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="56"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="58"
    End
    Begin
        dbText "Name" ="FABankBill.LineTot"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="65"
    End
    Begin
        dbText "Name" ="FABankBill.Bank_Assignee.ShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="68"
    End
    Begin
        dbText "Name" ="FABankBill.UnitCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="73"
    End
    Begin
        dbText "Name" ="RentNoTax"
        dbInteger "ColumnOrder" ="66"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbInteger "ColumnWidth" ="2205"
        dbInteger "ColumnOrder" ="67"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLTax"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.[FSLFee]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.FSLFee_EndDate.DateFeesEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.BillBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillToGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_Client_FAID_PO_Lease_Sch_Grp_Units.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillToSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Units.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BankContract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Schedule.FAVendorNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Schedule.FAPOnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemitAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankContract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateFeesEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToMBMAdr"
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =43
    Right =2087
    Bottom =878
    Left =-1
    Top =-1
    Right =2030
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =42
        Top =14
        Right =197
        Bottom =315
        Top =0
        Name ="FABankBill"
        Name =""
    End
End
