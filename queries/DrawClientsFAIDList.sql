SELECT
  Clients.clientID,
  Clients.ClientShNm AS Client,
  dbo_vw_Sch_Grp_Units_ChildParts.FAID,
  IIf(
    Count([drawdowns].[FAID])> 0,
    "Yes",
    "No"
  ) AS Drawn
FROM
  (
    Clients
    INNER JOIN dbo_vw_Sch_Grp_Units_ChildParts ON (
      Clients.clientID = dbo_vw_Sch_Grp_Units_ChildParts.clientID
    )
    AND (
      Clients.clientID = dbo_vw_Sch_Grp_Units_ChildParts.clientID
    )
  )
  LEFT JOIN DrawDowns ON dbo_vw_Sch_Grp_Units_ChildParts.FAID = DrawDowns.FAID
GROUP BY
  Clients.clientID,
  Clients.ClientShNm,
  dbo_vw_Sch_Grp_Units_ChildParts.FAID,
  Clients.clientType
HAVING
  (
    (
      (Clients.clientType)= "CLIENT"
    )
  )
ORDER BY
  Clients.clientID DESC,
  Clients.ClientShNm;
