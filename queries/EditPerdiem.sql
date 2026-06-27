SELECT
  Units.*,
  Nz(
    DMax(
      "UnitTitleLogDate", "UnitTitleLog",
      "EntryReasonId= 18 and UnitRef=" & [UnitID]
    ),
    ""
  ) AS tSent,
  Nz(
    DMax(
      "UnitTitleLogDate", "UnitTitleLog",
      "EntryReasonId= 1 and UnitRef=" & [UnitID]
    ),
    ""
  ) AS tBack
FROM
  Units;
