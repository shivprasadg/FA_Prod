SELECT
  DISTINCT vw_ATF_Avail_Signed_FAID.ClientGroupId,
  vw_ATF_Avail_Signed_FAID.clientID,
  Salesperson.SalespersonID AS BDEX,
  vw_ATF_Avail_Signed_FAID.clientID AS LesseeID,
  Clients.clientCompanyName AS Lessee,
  Clients.clientTaxID AS TaxID,
  vw_ATF_Avail_Signed_FAID.BDEX AS LOExec,
  vw_ATF_Avail_Signed_FAID.FAID,
  vw_ATF_Avail_Signed_FAID.SchName,
  vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc AS SchShortDesc,
  vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty AS DealUnitQty,
  vw_ATF_Avail_Signed_FAID.ShortGUID
FROM
  (
    vw_ATF_Avail_Signed_FAID
    LEFT JOIN Clients ON vw_ATF_Avail_Signed_FAID.clientID = Clients.clientID
  )
  LEFT JOIN Salesperson ON vw_ATF_Avail_Signed_FAID.BDEX = Salesperson.SPInitials
GROUP BY
  vw_ATF_Avail_Signed_FAID.ClientGroupId,
  vw_ATF_Avail_Signed_FAID.clientID,
  Salesperson.SalespersonID,
  vw_ATF_Avail_Signed_FAID.clientID,
  Clients.clientCompanyName,
  Clients.clientTaxID,
  vw_ATF_Avail_Signed_FAID.BDEX,
  vw_ATF_Avail_Signed_FAID.FAID,
  vw_ATF_Avail_Signed_FAID.SchName,
  vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc,
  vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty,
  vw_ATF_Avail_Signed_FAID.ShortGUID
ORDER BY
  Clients.clientCompanyName,
  vw_ATF_Avail_Signed_FAID.FAID DESC;
