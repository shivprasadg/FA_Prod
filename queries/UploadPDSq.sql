UPDATE
  Units
  INNER JOIN PDSUpload ON Units.UnitID = PDSUpload.[unit no]
SET
  Units.UnitPDStartDate = [pds];
