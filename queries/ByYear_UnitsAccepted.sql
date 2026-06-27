SELECT
  vw_sixkeys.ClientShortName,
  Min(
    DatePart("yyyy", [AcceptedDate])
  ) AS AYear,
  Count(vw_sixkeys.AssetID) AS UnitAccepted,
  vw_sixkeys.clientID,
  vw_sixkeys.ClientGroupId,
  Min(vw_sixkeys.AcceptedDate) AS FristAccepted,
  Max(vw_sixkeys.AcceptedDate) AS LastAccepted
FROM
  vw_sixkeys
WHERE
  (
    (
      (vw_sixkeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_sixkeys.ClientShortName,
  vw_sixkeys.clientID,
  vw_sixkeys.ClientGroupId
ORDER BY
  Min(
    DatePart("yyyy", [AcceptedDate])
  ) DESC;
