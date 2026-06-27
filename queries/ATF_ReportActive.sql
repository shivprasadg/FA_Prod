SELECT
  vw_ATF_Active_Profile.SchID,
  vw_ATF_Active_Profile.ATFID,
  vw_ATF_Active_Profile.FAID,
  vw_ATF_Active_Profile.MLID,
  vw_ATF_Active_Profile.LesseeID,
  vw_ATF_Active_Profile.CoLesseeID,
  vw_ATF_Active_Profile.SubLeasee,
  vw_ATF_Active_Profile.CreditRank,
  vw_ATF_Active_Profile.ATFSchSpecID,
  Max(
    vw_ATF_Active_Profile.Assignee_Bank
  ) AS Assignee_Bank,
  vw_ATF_Active_Profile.SchFAID,
  vw_ATF_Active_Profile.CompanyName,
  vw_ATF_Active_Profile.ClientShortName,
  vw_ATF_Active_Profile.clientTaxID,
  vw_ATF_Active_Profile.MLNo,
  vw_ATF_Active_Profile.MLOrig,
  vw_ATF_Active_Profile.MLDate,
  Sum(
    vw_ATF_Active_Profile.ActualUnits
  ) AS ActualUnits,
  Sum(
    vw_ATF_Active_Profile.UnitsCost
  ) AS SumOfUnitsCost,
  Sum(vw_ATF_Active_Profile.Parts) AS SumOfParts,
  Sum(
    vw_ATF_Active_Profile.PartsCost
  ) AS SumOfPartsCost,
  Sum(vw_ATF_Active_Profile.UnitOEC) AS SumOfUnitOEC,
  Avg(
    vw_ATF_Active_Profile.SchRentAvg
  ) AS AvgOfSchRentAvg,
  Avg(
    vw_ATF_Active_Profile.SyndRentAvg
  ) AS AvgOfSyndRentAvg,
  Avg(
    vw_ATF_Active_Profile.ExtendedRentAvg
  ) AS AvgOfExtendedRentAvg,
  Avg(
    vw_ATF_Active_Profile.MTMRentAvg
  ) AS AvgOfMTMRentAvg,
  Avg(
    vw_ATF_Active_Profile.PDRentAvg
  ) AS AvgOfPDRentAvg,
  Avg(
    vw_ATF_Active_Profile.GroupLeasePayment
  ) AS AvgOfGroupLeasePayment,
  vw_ATF_Active_Profile.MLNo AS ATFDesc,
  Ecount(
    "SchID", "vw_SixKeys", "SchID=" & [SchID] & "",
    True
  ) AS NumOfSchedules,
  Ecount(
    "sGrpID", "vw_SixKeys", "SchID=" & [SchID] & "",
    True
  ) AS NumOfGroups,
  Ecount(
    "UnitID", "vw_SixKeys", "SchID=" & [SchID] & "",
    True
  ) AS NumOfUnits,
  Null AS CoLessee,
  vw_ATF_Active_Profile.ATFDate,
  vw_ATF_Active_Profile.LPAFDate,
  vw_ATF_Active_Profile.BDEX,
  vw_ATF_Active_Profile.SchDescription,
  vw_ATF_Active_Profile.SchNo,
  vw_ATF_Active_Profile.Assigned,
  DLookUp(
    "SFDCType", "Deals", "FAID='" & [FAID] & "'"
  ) AS SFDCType,
  Count(vw_ATF_Active_Profile.sGrpID) AS GroupsIn,
  vw_ATF_Active_Profile.SGrpID,
  vw_ATF_Active_Profile.SGrpID AS GroupID,
  vw_ATF_Active_Profile.GroupNo
FROM
  vw_ATF_Active_Profile
GROUP BY
  vw_ATF_Active_Profile.SchID,
  vw_ATF_Active_Profile.ATFID,
  vw_ATF_Active_Profile.FAID,
  vw_ATF_Active_Profile.MLID,
  vw_ATF_Active_Profile.LesseeID,
  vw_ATF_Active_Profile.CoLesseeID,
  vw_ATF_Active_Profile.SubLeasee,
  vw_ATF_Active_Profile.CreditRank,
  vw_ATF_Active_Profile.ATFSchSpecID,
  vw_ATF_Active_Profile.SchFAID,
  vw_ATF_Active_Profile.CompanyName,
  vw_ATF_Active_Profile.ClientShortName,
  vw_ATF_Active_Profile.clientTaxID,
  vw_ATF_Active_Profile.MLOrig,
  vw_ATF_Active_Profile.MLDate,
  vw_ATF_Active_Profile.MLNo,
  vw_ATF_Active_Profile.ATFDate,
  vw_ATF_Active_Profile.LPAFDate,
  vw_ATF_Active_Profile.BDEX,
  vw_ATF_Active_Profile.SchDescription,
  vw_ATF_Active_Profile.SchNo,
  vw_ATF_Active_Profile.Assigned,
  vw_ATF_Active_Profile.SGrpID,
  vw_ATF_Active_Profile.MLNo,
  vw_ATF_Active_Profile.SGrpID,
  vw_ATF_Active_Profile.GroupNo;
