SELECT
  Units.UnitID,
  LastTitleEntry_In.TitleNumber,
  LastTitleEntry_In.LastOfUnitTitleLogDate AS DateLog INTO TitleLogTemp
FROM
  Units
  INNER JOIN LastTitleEntry_In ON Units.UnitID = LastTitleEntry_In.UnitRef
WHERE
  (
    (
      (LastTitleEntry_In.TitleNumber) Is Not Null
    )
    AND (
      (Units.UnitTitleNumber) Is Null
    )
  );
