SELECT
  DrawDowns.DrawDownId,
  Banks.BankShortName AS Bank,
  DrawDowns.DrawNumber,
  DrawDowns.DrawBankAccountNumber,
  DrawDowns.DrawRequestNumber,
  Units.UnitID AS AssetID,
  Units.UnitVIN AS VIN,
  Units.UnitUnitNum AS UnitNumber,
  Units.UnitStatus,
  Units.UnitActualDelDate AS ActualDelivery,
  Units.UnitInvoiceNum AS InvoiceNum,
  Units.UnitInvoiceDate AS InvoiceDate,
  Units.UnitVendorPytDate AS VendorPaid,
  VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName,
  VW_Clients_ML_FAID_SCH_Groups_count.MLNo,
  VW_Clients_ML_FAID_SCH_Groups_count.FAID,
  VW_Clients_ML_FAID_SCH_Groups_count.Schedule,
  VW_Clients_ML_FAID_SCH_Groups_count.UnitGroup,
  DrawDowns.ClientID,
  VW_Clients_ML_FAID_SCH_Groups_count.SchID,
  VW_Clients_ML_FAID_SCH_Groups_count.SGrpID,
  Units.VendorREF AS VendorID
FROM
  (
    (
      DrawDowns
      INNER JOIN Units ON DrawDowns.DrawDownId = Units.DrawDownId
    )
    INNER JOIN Banks ON DrawDowns.DrawBankId = Banks.BankId
  )
  INNER JOIN VW_Clients_ML_FAID_SCH_Groups_count ON Units.GroupREF = VW_Clients_ML_FAID_SCH_Groups_count.SGrpID;
