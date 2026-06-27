SELECT
  UnitLevelTaxes.UnitID,
  UnitLevelTaxes.UnitTaxes
FROM
  UnitLevelTaxes
  INNER JOIN Bulk_Update_WorkTable ON UnitLevelTaxes.UnitID = Bulk_Update_WorkTable.UnitID;
