SELECT
  vw_Sch_Grp_Units_Remarketing.clientID,
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  Nz(
    [vw_Sch_Grp_Units_Remarketing].[ClientShortName] & "_RMKT"
  ) AS FAID
FROM
  vw_Sch_Grp_Units_Remarketing
GROUP BY
  vw_Sch_Grp_Units_Remarketing.clientID,
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  Nz(
    [vw_Sch_Grp_Units_Remarketing].[ClientShortName] & "_RMKT"
  );
