SELECT
  Nz(
    [vw_Sch_Grp_Units_Remarketing].[faid],
    [ClientShortName] & "_RMKT"
  ) AS FAID,
  vw_Sch_Grp_Units_Remarketing.Schedule,
  vw_Sch_Grp_Units_Remarketing.UnitGroup AS [Group],
  Count(
    vw_Sch_Grp_Units_Remarketing.AssetId
  ) AS Units,
  vw_Sch_Grp_Units_Remarketing.SGrpID,
  vw_Sch_Grp_Units_Remarketing.SchID,
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  vw_Sch_Grp_Units_Remarketing.clientID
FROM
  vw_Sch_Grp_Units_Remarketing
GROUP BY
  Nz(
    [vw_Sch_Grp_Units_Remarketing].[faid],
    [ClientShortName] & "_RMKT"
  ),
  vw_Sch_Grp_Units_Remarketing.Schedule,
  vw_Sch_Grp_Units_Remarketing.UnitGroup,
  vw_Sch_Grp_Units_Remarketing.SGrpID,
  vw_Sch_Grp_Units_Remarketing.SchID,
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  vw_Sch_Grp_Units_Remarketing.clientID
ORDER BY
  vw_Sch_Grp_Units_Remarketing.Schedule,
  vw_Sch_Grp_Units_Remarketing.UnitGroup;
