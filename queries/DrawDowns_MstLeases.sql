SELECT
  vw_SixKeys.MLID,
  vw_SixKeys.MLNo,
  Count(vw_SixKeys.AssetID) AS Units,
  Max(
    IIf(
      Nz([vw_SixKeys].[DrawDownId], 0)<> 0,
      'Yes',
      'No'
    )
  ) AS [Has Draw],
  vw_SixKeys.ClientID
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.MLID,
  vw_SixKeys.MLNo,
  vw_SixKeys.ClientID
ORDER BY
  vw_SixKeys.MLNo;
