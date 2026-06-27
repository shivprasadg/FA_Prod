SELECT
  [costco repairsNew].Unitref,
  Max([costco repairsNew].HubMI) AS MaxOfHubMI,
  Max(
    [costco repairsNew].[Close Date]
  ) AS [MaxOfClose Date],
  SchGrp.SGrpID,
  SchGrp.UnitsInGroup,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1 AS Expiry,
  DateDiff(
    "d",
    [BLCD],
    (
      DateAdd(
        "m", [LeaseTermPrimary], [BLCD] - 1
      )
    )
  ) AS Days,
  ([LeaseTermPrimary] / 12)* [SgprAllow1] AS LTDMI,
  Units.unitinserv
FROM
  SchGrp
  INNER JOIN (
    [costco repairsNew]
    INNER JOIN Units ON [costco repairsNew].Unitref = Units.UnitID
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  [costco repairsNew].Unitref,
  SchGrp.SGrpID,
  SchGrp.UnitsInGroup,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1,
  DateDiff(
    "d",
    [BLCD],
    (
      DateAdd(
        "m", [LeaseTermPrimary], [BLCD] - 1
      )
    )
  ),
  ([LeaseTermPrimary] / 12)* [SgprAllow1],
  Units.unitinserv
HAVING
  (
    (
      (
        Max([costco repairsNew].HubMI)
      )> 0
    )
  )
ORDER BY
  [costco repairsNew].Unitref;
