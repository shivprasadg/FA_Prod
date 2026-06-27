SELECT
  DISTINCTROW RemarketingUnitsBySchedule.UnitID,
  RemarketingUnitsBySchedule.Unitunitnum AS UnitNum,
  RemarketingUnitsBySchedule.Unitvin,
  RemarketingUnitsBySchedule.SchID,
  RemarketingUnitsBySchedule.SchNo,
  RemarketingUnitsBySchedule.UnitStatus,
  RemarketingUnitsBySchedule.BuyOutDate,
  RemarketingUnitsBySchedule.BuyOutCost,
  RemarketingUnitsBySchedule.IFBankID,
  RemarketingUnitsBySchedule.IFBank,
  RemarketingUnitsBySchedule.IFLoanNumber,
  RemarketingUnitsBySchedule.IFDate,
  RemarketingUnitsBySchedule.IFAmount,
  RemarketingUnitsBySchedule.IFPaidDate,
  RemarketingUnitsBySchedule.InvoiceSentDate,
  RemarketingUnitsBySchedule.InvoicePaidDate,
  RemarketingUnitsBySchedule.SoldPrice,
  RemarketingUnitsBySchedule.BuyDateEffective
FROM
  ReMarketingBuyOut_WorkTable
  INNER JOIN RemarketingUnitsBySchedule ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID;
