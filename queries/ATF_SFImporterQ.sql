SELECT
  vw_ATF_Avail_Signed_FAID.clientID,
  vw_ATF_Avail_Signed_FAID.Clients_clientCompanyName,
  vw_ATF_Avail_Signed_FAID.FAID,
  vw_ATF_Avail_Signed_FAID.BDEX,
  vw_ATF_Avail_Signed_FAID.Schedules AS Schedules,
  IIf(
    [Schedules] > 1,
    Max([GroupsIn])* [Schedules],
    Max([GroupsIn])
  ) AS Groups,
  First(
    vw_ATF_Avail_Signed_FAID.ShortGUID
  ) AS ShortGUID,
  First(
    vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc
  ) AS Schedule_SchShortDesc,
  Sum(
    vw_ATF_Avail_Signed_FAID.Spec_UnitQty
  ) AS Spec_UnitQty,
  vw_ATF_Avail_Signed_FAID.SalesPersonID,
  vw_ATF_Avail_Signed_FAID.BDEX AS SPInitials,
  vw_ATF_Avail_Signed_FAID.ClientGroupId,
  "2/5/2019" AS LPAFcomp,
  vw_ATF_Avail_Signed_FAID.GuarantorName,
  vw_ATF_Avail_Signed_FAID.GuarantorID,
  vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc,
  vw_ATF_Avail_Signed_FAID.SchGrp_OrigEquipCost,
  vw_ATF_Avail_Signed_FAID.UnitDesc,
  vw_ATF_Avail_Signed_FAID.[LPAFDate] AS ActualClose,
  vw_ATF_Avail_Signed_FAID.SF_ClientLegalName,
  vw_ATF_Avail_Signed_FAID.Spec_Name,
  vw_ATF_Avail_Signed_FAID.Deals_dealSchedNum,
  vw_ATF_Avail_Signed_FAID.LPAFDate AS LPAFDate,
  vw_ATF_Avail_Signed_FAID.SFDC_Accounts_Id
FROM
  vw_ATF_Avail_Signed_FAID
GROUP BY
  vw_ATF_Avail_Signed_FAID.clientID,
  vw_ATF_Avail_Signed_FAID.Clients_clientCompanyName,
  vw_ATF_Avail_Signed_FAID.FAID,
  vw_ATF_Avail_Signed_FAID.Schedules,
  vw_ATF_Avail_Signed_FAID.SalesPersonID,
  vw_ATF_Avail_Signed_FAID.BDEX,
  vw_ATF_Avail_Signed_FAID.ClientGroupId,
  vw_ATF_Avail_Signed_FAID.GuarantorName,
  vw_ATF_Avail_Signed_FAID.GuarantorID,
  vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc,
  vw_ATF_Avail_Signed_FAID.SchGrp_OrigEquipCost,
  vw_ATF_Avail_Signed_FAID.UnitDesc,
  vw_ATF_Avail_Signed_FAID.[LPAFDate],
  vw_ATF_Avail_Signed_FAID.SF_ClientLegalName,
  vw_ATF_Avail_Signed_FAID.Spec_Name,
  vw_ATF_Avail_Signed_FAID.Deals_dealSchedNum,
  vw_ATF_Avail_Signed_FAID.LPAFDate,
  vw_ATF_Avail_Signed_FAID.SFDC_Accounts_Id
ORDER BY
  vw_ATF_Avail_Signed_FAID.SF_ClientLegalName;
