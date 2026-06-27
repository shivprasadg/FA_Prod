UPDATE
  DISTINCTROW TranscoInservice
  INNER JOIN Units ON TranscoInservice.UnitNum = Units.UnitUnitNum
SET
  Units.UnitInServ = [TranscoInservice].[inservicedate],
  Units.ReportStartDate = [TranscoInservice].[inservicedate],
  Units.UnitStatus = "A",
  Units.UnitInServInd = "A"
WHERE
  (
    (
      (Units.UnitInServ) Is Null
    )
    AND (
      (TranscoInservice.InServiceDate) Is Not Null
    )
  );
