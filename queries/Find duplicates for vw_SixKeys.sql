SELECT
  vw_SixKeys.VIN,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.SGrpID,
  vw_SixKeys.CompanyName
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.VIN) In (
        SELECT
          [VIN]
        FROM
          [vw_SixKeys] As Tmp
        GROUP BY
          [VIN]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  vw_SixKeys.VIN;
