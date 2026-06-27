SELECT
  vw_SixKeys.MLOrig,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.UnitID,
  IIf(
    [vw_SixKeys].[BLCD] < #1/1/2015#
    And Nz([unitoffleasedt], #1/1/2030#)> #12/31/2015#,
    1,
    0
  ) AS OnLease2015,
  IIf(
    [BLCD] Between #1/1/2015#
    And #12/31/2015#,
    1, 0
  ) AS Added2015,
  IIf(
    [unitoffleasedt] Between #1/1/2015#
    And #12/31/2015#,
    1, 0
  ) AS Terminated2015,
  IIf(
    [BLCD] Between #1/1/2016#
    And #12/31/2016#,
    1, 0
  ) AS Added2016,
  IIf(
    [unitoffleasedt] Between #1/1/2016#
    And #12/31/2016#,
    1, 0
  ) AS Terminated2016,
  IIf(
    DateDiff(
      "m",
      [ProjEOL],
      Nz(
        [unitoffleasedt],
        Date()
      )
    )>= 1,
    1,
    0
  ) AS Extended,
  IIf(
    [OnLease2015] = 1
    And [Extended] = 1,
    1, 0
  ) AS Extended2015,
  IIf(
    [Added2016] = 1
    And [Extended] = 1, 1,
    0
  ) AS Extended2016,
  DateDiff(
    "m",
    [ProjEOL],
    Nz(
      [unitoffleasedt],
      Date()
    )
  ) AS MonthsExt,
  vw_SixKeys.ProjEOL,
  vw_SixKeys.OffLeaseDate,
  DateAdd(
    "m", [LeaseTermPrimary], [BLCD] - 1
  ) AS CalcEOL
FROM
  vw_SixKeys
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
GROUP BY
  vw_SixKeys.MLOrig,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.UnitID,
  IIf(
    [vw_SixKeys].[BLCD] < #1/1/2015#
    And Nz([unitoffleasedt], #1/1/2030#)> #12/31/2015#,
    1,
    0
  ),
  IIf(
    [BLCD] Between #1/1/2015#
    And #12/31/2015#,
    1, 0
  ),
  IIf(
    [unitoffleasedt] Between #1/1/2015#
    And #12/31/2015#,
    1, 0
  ),
  IIf(
    [BLCD] Between #1/1/2016#
    And #12/31/2016#,
    1, 0
  ),
  IIf(
    [unitoffleasedt] Between #1/1/2016#
    And #12/31/2016#,
    1, 0
  ),
  IIf(
    DateDiff(
      "m",
      [ProjEOL],
      Nz(
        [unitoffleasedt],
        Date()
      )
    )>= 1,
    1,
    0
  ),
  DateDiff(
    "m",
    [ProjEOL],
    Nz(
      [unitoffleasedt],
      Date()
    )
  ),
  vw_SixKeys.ProjEOL,
  vw_SixKeys.OffLeaseDate,
  DateAdd(
    "m", [LeaseTermPrimary], [BLCD] - 1
  )
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (
        IIf(
          DateDiff(
            "m",
            [ProjEOL],
            Nz(
              [unitoffleasedt],
              Date()
            )
          )>= 1,
          1,
          0
        )
      )= 0
    )
    AND (
      (vw_SixKeys.OffLeaseDate) Between #1/1/2016#
      And #12/31/2016#
    )
  );
