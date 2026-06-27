SELECT
  vw_Sch_Grp_Units_Remarketing.clientID,
  vw_Sch_Grp_Units_Remarketing.CompanyName,
  vw_Sch_Grp_Units_Remarketing.AssetID,
  vw_Sch_Grp_Units_Remarketing.VIN,
  Rmkt.BuyOutCost,
  Rmkt.BuyOutDate,
  Rmkt.[SUF Sent],
  Rmkt.[SUF Returned],
  vw_Sch_Grp_Units_Remarketing.Status
FROM
  vw_Sch_Grp_Units_Remarketing
  INNER JOIN Rmkt ON vw_Sch_Grp_Units_Remarketing.AssetID = Rmkt.UnitRef
WHERE
  (
    (
      (Rmkt.[SUF Sent]) Is Null
    )
  )
ORDER BY
  vw_Sch_Grp_Units_Remarketing.AssetID;
