SELECT
  Clients.clientID,
  Clients.ClientShNm AS Client,
  IIf(
    Count(
      [drawdowns].[DrawRequestNumber]
    )> 0,
    "Yes",
    "No"
  ) AS Draws
FROM
  (
    Clients
    INNER JOIN vw_Sch_Grp_Units_ChildParts ON (
      Clients.clientID = vw_Sch_Grp_Units_ChildParts.clientID
    )
    AND (
      Clients.clientID = vw_Sch_Grp_Units_ChildParts.clientID
    )
  )
  LEFT JOIN DrawDowns ON vw_Sch_Grp_Units_ChildParts.MLID = DrawDowns.MLID
GROUP BY
  Clients.clientID,
  Clients.ClientShNm,
  Clients.clientType
HAVING
  (
    (
      (Clients.clientType)= "CLIENT"
    )
  )
ORDER BY
  Clients.ClientShNm;
