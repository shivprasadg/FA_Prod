UPDATE
  Units
SET
  Units.UnitStatus = "O",
  Units.UnitOLInd = "A",
  Units.ReportEndDate = [unitoffleasedt]
WHERE
  (
    (
      (Units.UnitStatus)= "T"
    )
    AND (
      (Units.unitoffleasedt) Is Not Null
      And (Units.unitoffleasedt)< Date()-30
    )
  );
