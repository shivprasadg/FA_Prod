SELECT
  Units.UnitID,
  UnitTitleLog.UnitTitleLogDate,
  IIf([entrytypeid] = 1, "In", "Out") AS Status,
  Reasons.ReasonDescription,
  UnitTitleLog.TitleNumber,
  Units.unitvin
FROM
  (
    Units
    INNER JOIN UnitTitleLog ON Units.UnitID = UnitTitleLog.UnitRef
  )
  INNER JOIN Reasons ON UnitTitleLog.EntryReasonId = Reasons.ReasonId
WHERE
  (
    (
      (Units.unitvin)= [forms]![titlecheckin]![seekvin]
    )
  )
ORDER BY
  UnitTitleLog.UnitTitleLogDate DESC;
